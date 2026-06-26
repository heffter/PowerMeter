#!/usr/bin/env python3
"""
calibrate_tvn.py -- Batch RFG sensor and drive calibration uploader.

Reads 8 auto-calibration log files produced by Remote_Control auto-cal sweep
(one file per RFG/channel/direction), fits sensor polynomials (FORWARD/REFLECTED)
per frequency range and a per-channel drive polynomial (POWER), then uploads all
CALIBRATE commands to both RFG boards over the single FSD serial bus used by
Remote_Terminal, and writes calibrate_rfg_0.csv / calibrate_rfg_1.csv as
side-effect files for record-keeping.

Replaces the Excel-based workflow (automate_calibration.bat + automate_excel.ps1)
and eliminates the manual Remote Terminal upload step.

Usage:
    python calibrate_tvn.py --dir C:\\cal_data\\TVN-42 --port COM60
    python calibrate_tvn.py --dir C:\\cal_data\\TVN-42           # dry run (print only)
    python calibrate_tvn.py --dir C:\\cal_data\\TVN-42 --port COM60 --rfg 0  # RFG0 only

Expected input files in --dir (log or csv extension, produced by Remote_Control):
    rfg_0AF*.{log,csv}  RFG0 channel A forward relay configuration
    rfg_0AR*.{log,csv}  RFG0 channel A reflected relay configuration
    rfg_0BF*.{log,csv}  RFG0 channel B forward relay configuration
    rfg_0BR*.{log,csv}  RFG0 channel B reflected relay configuration
    rfg_1AF*.{log,csv}  RFG1 channel A forward relay configuration
    rfg_1AR*.{log,csv}  RFG1 channel A reflected relay configuration
    rfg_1BF*.{log,csv}  RFG1 channel B forward relay configuration
    rfg_1BR*.{log,csv}  RFG1 channel B reflected relay configuration

The log file header format (Remote_Control 8-column CSV):
    Channel, Frequency, PWR - Control - Level, AMS - FWD - PWR, Bird - FWD - PWR,
    AMS - REF - PWR, Bird - REF - PWR, Drive-Sum

Sensor polynomial model (per frequency TABLE slot, per channel):
    bird_watts = a + b*x + c*x^2 + d*x^3 + e*x^4
    where x = sqrt(ams_watts)   -- the raw ADC input domain

Drive polynomial model (bivariate, stored in TABLE 0, per channel):
    volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)
    where f = freq_MHz, P = bird_fwd_watts, volt = sqrt(drive_sum) * DRIVE_VOLT_SCALE

Firmware sentinel: drive polynomial only activates when 0.01 < c < 10.0.

FSD protocol:
    Commands are wrapped in FSD TEXT (TT_KEYBOARD) packets at 921600 baud, the same
    protocol Remote_Terminal uses.  One serial port addresses both RFG boards via
    destination_id in the packet header (RF_GENERATOR_ID + board_number).

Prerequisites:
    pip install numpy pyserial
"""

import argparse
import glob
import math
import os
import struct
import sys
import time

import numpy as np

# ---------------------------------------------------------------------------
# FSD protocol constants  (from FSD_Packets.cs)
# ---------------------------------------------------------------------------

FP_SYNCH            = 0x24          # '$'
FP_TEXT             = 0x20
FP_TT_KEYBOARD      = 0x05
FP_RF_GENERATOR_ID  = 0x00002400    # base device ID for RFG boards
FP_THERAVISION_ID   = 0x00001100    # source ID for Remote Terminal / this tool
FP_PKT_HEADER_SIZE  = 16
FP_BAUD             = 921600

# unique_id inside the TEXT packet payload is a hardware-discovered per-board
# serial number used by Remote_Terminal for logging.  Routing is performed by
# destination_id in the packet header, so 0 is safe here.
FP_UNIQUE_ID        = 0

# ---------------------------------------------------------------------------
# Calibration constants
# ---------------------------------------------------------------------------

DRIVE_VOLT_SCALE    = 2.5 / math.sqrt(60.0)   # must match firmware rf_generator.hc
SENTINEL_C_MIN      = 0.01
SENTINEL_C_MAX      = 10.0
MAX_CALIBRATION_TABLES = 16

CMD_DELAY_S         = 0.7   # seconds between packets (matches Remote_Terminal 700 ms)

# Convergence quality gate applied per-frequency before fitting.
# A row passes when ams_fwd >= CONVERGENCE_THRESHOLD * pwr_setpoint.
# A frequency slot is included only when MIN_CONVERGED_ROWS rows pass.
CONVERGENCE_THRESHOLD = 0.80
MIN_CONVERGED_ROWS    = 4

# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

_FILE_PREFIXES = {
    (0, 0, 'F'): 'rfg_0AF',
    (0, 0, 'R'): 'rfg_0AR',
    (0, 1, 'F'): 'rfg_0BF',
    (0, 1, 'R'): 'rfg_0BR',
    (1, 0, 'F'): 'rfg_1AF',
    (1, 0, 'R'): 'rfg_1AR',
    (1, 1, 'F'): 'rfg_1BF',
    (1, 1, 'R'): 'rfg_1BR',
}


def find_log(directory, prefix):
    for ext in ('log', 'csv'):
        matches = glob.glob(os.path.join(directory, f'{prefix}*.{ext}'))
        if matches:
            return sorted(matches)[0]
    return None


# ---------------------------------------------------------------------------
# Log file parsing
# ---------------------------------------------------------------------------

def load_log(path):
    """
    Parse a Remote_Control auto-calibration log file.
    Returns list of row dicts with keys:
        channel, freq_khz, pwr_setpoint, ams_fwd, bird_fwd, ams_ref, bird_ref, drive_sum
    """
    rows = []
    with open(path, newline='', encoding='utf-8', errors='replace') as fh:
        raw_lines = fh.readlines()

    data_start = 1
    for idx, line in enumerate(raw_lines):
        if 'Channel' in line and 'Frequency' in line:
            data_start = idx + 1
            break

    for line in raw_lines[data_start:]:
        parts = [p.strip() for p in line.split(',')]
        if len(parts) < 7:
            continue
        try:
            rows.append({
                'channel':      int(parts[0]),
                'freq_khz':     int(parts[1]),
                'pwr_setpoint': float(parts[2]),
                'ams_fwd':      float(parts[3]),
                'bird_fwd':     float(parts[4]),
                'ams_ref':      float(parts[5]),
                'bird_ref':     float(parts[6]),
                'drive_sum':    float(parts[7]) if len(parts) > 7 else 0.0,
            })
        except (ValueError, IndexError):
            continue

    return rows


# ---------------------------------------------------------------------------
# Convergence quality gate
# ---------------------------------------------------------------------------

def valid_frequencies(fwd_rows, threshold=CONVERGENCE_THRESHOLD, min_rows=MIN_CONVERGED_ROWS):
    """
    Return (valid_set, excluded_dict) from forward-relay rows.

    A row converges when ams_fwd >= threshold * pwr_setpoint.
    A frequency is included when it has >= min_rows converging rows.
    excluded_dict maps freq_khz -> (converged_count, total_count).
    """
    counts = {}
    for row in fwd_rows:
        f = row['freq_khz']
        if f not in counts:
            counts[f] = [0, 0]
        sp = row['pwr_setpoint']
        counts[f][0] += 1
        if sp <= 0 or row['ams_fwd'] >= threshold * sp:
            counts[f][1] += 1

    valid    = {f for f, (tot, conv) in counts.items() if conv >= min_rows}
    excluded = {f: (conv, tot) for f, (tot, conv) in counts.items() if conv < min_rows}
    return valid, excluded


# ---------------------------------------------------------------------------
# Polynomial fitting
# ---------------------------------------------------------------------------

_IDENTITY_POLY = (0.0, 0.0, 1.0, 0.0, 0.0)


def fit_sensor_poly(ams_values, bird_values):
    """
    Fit 4th-degree polynomial: bird = a + b*x + c*x^2 + d*x^3 + e*x^4
    where x = sqrt(ams).

    The identity polynomial (a=0, b=0, c=1, d=0, e=0) gives bird = ams,
    i.e. no correction.  Returns identity when data is too sparse.
    """
    valid = [(a, b) for a, b in zip(ams_values, bird_values) if a >= 0 and b > 0]
    if len(valid) < 3:
        return _IDENTITY_POLY
    x = np.array([math.sqrt(a) for a, _ in valid])
    y = np.array([b for _, b in valid])
    # np.polyfit returns descending order: p = [e, d, c, b, a] for degree 4
    p = np.polyfit(x, y, deg=4)
    return (float(p[4]), float(p[3]), float(p[2]), float(p[1]), float(p[0]))


def fit_drive_poly(rows, min_power=1.0):
    """
    Fit bivariate drive polynomial from forward-relay measurements.
    Model: volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)
    Returns (a, b, c, d, e) or None when data is insufficient.
    """
    pts = []
    for row in rows:
        if row['drive_sum'] <= 0.0 or row['bird_fwd'] < min_power:
            continue
        f    = row['freq_khz'] / 1000.0
        p    = row['bird_fwd']
        volt = math.sqrt(row['drive_sum']) * DRIVE_VOLT_SCALE
        pts.append((f, p, volt))
    if len(pts) < 6:
        return None
    A = np.array([[1.0, f, math.sqrt(p), f * f, f * math.sqrt(p)] for f, p, _ in pts])
    y = np.array([v for _, _, v in pts])
    coeffs, _, _, _ = np.linalg.lstsq(A, y, rcond=None)
    return tuple(float(c) for c in coeffs)


# ---------------------------------------------------------------------------
# Frequency table structure
# ---------------------------------------------------------------------------

def build_tables(fwd0_rows, ref0_rows, fwd1_rows, ref1_rows, valid_freqs=None):
    """
    Group measurements by frequency and fit per-slot polynomials.
    Frequencies absent from valid_freqs (when supplied) are skipped.
    Returns list of table dicts sorted ascending by frequency.
    """
    all_rows = fwd0_rows + ref0_rows + fwd1_rows + ref1_rows
    all_freqs = sorted(set(r['freq_khz'] for r in all_rows))
    freqs = [f for f in all_freqs if valid_freqs is None or f in valid_freqs]
    if not freqs:
        return []

    tables = []
    n = len(freqs)
    for i, freq in enumerate(freqs):
        half_prev = (freq - freqs[i - 1]) // 2 if i > 0 else ((freqs[1] - freq) // 2 if n > 1 else 500)
        half_next = (freqs[i + 1] - freq) // 2 if i < n - 1 else ((freq - freqs[i - 1]) // 2 if n > 1 else 500)
        f_lo = max(0, freq - half_prev)
        f_hi = freq + half_next

        def at(rows, f=freq):
            return [r for r in rows if r['freq_khz'] == f]

        f0 = at(fwd0_rows)
        r0 = at(ref0_rows)
        f1 = at(fwd1_rows)
        r1 = at(ref1_rows)

        fwd_vals = [r['bird_fwd'] for r in f0 + f1 if r['bird_fwd'] > 0]
        ref_vals = [r['bird_ref'] for r in r0 + r1 if r['bird_ref'] > 0]
        fwd_max  = max(1, math.ceil(max(fwd_vals))) if fwd_vals else 35
        ref_max  = max(1, math.ceil(max(ref_vals))) if ref_vals else 10

        tables.append({
            'freq_khz': freq,
            'freq_lo':  f_lo,
            'freq_hi':  f_hi,
            'fwd_max':  fwd_max,
            'ref_max':  ref_max,
            'fwd0': fit_sensor_poly([r['ams_fwd'] for r in f0], [r['bird_fwd'] for r in f0]),
            'ref0': fit_sensor_poly([r['ams_ref'] for r in r0], [r['bird_ref'] for r in r0]),
            'fwd1': fit_sensor_poly([r['ams_fwd'] for r in f1], [r['bird_fwd'] for r in f1]),
            'ref1': fit_sensor_poly([r['ams_ref'] for r in r1], [r['bird_ref'] for r in r1]),
        })

    return tables


# ---------------------------------------------------------------------------
# Command generation
# ---------------------------------------------------------------------------

def _coeff_str(a, b, c, d, e):
    # Use %.9g: IEEE-754 float32 needs 9 significant digits for exact round-trip.
    # %.8f gave only ~5 sig-figs on small coefficients (e.g. 0.00038667 -> ~2.6% error).
    # atof() on the RFG firmware parses scientific notation (e.g. 3.86670e-04) correctly.
    return f'{a:.9g} {b:.9g} {c:.9g} {d:.9g} {e:.9g}'


def build_commands(tables, drive_ch0, drive_ch1):
    """
    Generate the complete CALIBRATE ASCII command sequence for one RFG.
    Drive polynomial is stored in TABLE 0 alongside the first sensor slot.
    Returns list of ASCII command strings (no line endings).
    """
    cmds = []
    for slot, tbl in enumerate(tables):
        cmds.append(f'CALIBRATE TABLE {slot}')
        for ch in ('1', '2'):
            cmds.append(
                f'CALIBRATE {ch} RANGE '
                f'{tbl["freq_lo"]} {tbl["freq_hi"]} '
                f'0 {tbl["fwd_max"]} '
                f'0 {tbl["ref_max"]} '
                f'0 {tbl["fwd_max"]}'
            )
        cmds.append(f'CALIBRATE 1 FORWARD {_coeff_str(*tbl["fwd0"])}')
        cmds.append(f'CALIBRATE 2 FORWARD {_coeff_str(*tbl["fwd1"])}')
        cmds.append(f'CALIBRATE 1 REFLECTED {_coeff_str(*tbl["ref0"])}')
        cmds.append(f'CALIBRATE 2 REFLECTED {_coeff_str(*tbl["ref1"])}')
        if slot == 0:
            if drive_ch0 is not None:
                cmds.append(f'CALIBRATE 1 POWER {_coeff_str(*drive_ch0)}')
            if drive_ch1 is not None:
                cmds.append(f'CALIBRATE 2 POWER {_coeff_str(*drive_ch1)}')
        cmds.append('CALIBRATE WRITE')
    return cmds


# ---------------------------------------------------------------------------
# FSD packet builder
# ---------------------------------------------------------------------------

def _crc16_fsd(data: bytes) -> int:
    """
    CRC-16/CCITT as implemented in FSD crc.cs.
    Initial value 0xFFFF, polynomial 0x1021.
    Each input byte is shifted into the MSB of a 16-bit accumulator and
    processed bit-by-bit.
    """
    crc = 0xFFFF
    for byte_val in data:
        d = byte_val << 8
        for _ in range(8):
            if ((d ^ crc) & 0x8000) != 0:
                crc = ((crc << 1) ^ 0x1021) & 0xFFFF
            else:
                crc = (crc << 1) & 0xFFFF
            d = (d << 1) & 0xFFFF
    return crc


def build_fsd_text_packet(destination_id: int, command: str) -> bytes:
    """
    Build one FSD TEXT (TT_KEYBOARD) packet wrapping an ASCII CALIBRATE command.

    Packet layout:
        Header (16 bytes, all big-endian):
            synch           1  0x24
            packet_type     1  0x20 (TEXT)
            packet_size     2  total length of this packet in bytes
            destination_id  4  FP_RF_GENERATOR_ID + board_number
            source_id       4  FP_THERAVISION_ID
            sequence_num    1  0
            broadcast_dst   1  0
            broadcast_depth 1  0
            repeat_count    1  0  (also forced to 0 during CRC calculation)
        Payload:
            unique_id       4  0 (routing uses destination_id, not unique_id)
            tt_type         1  0x05 (TT_KEYBOARD)
            data_size       2  len(command) + 1 (for the trailing newline)
            data            N  command bytes + 0x0A newline
        CRC-16             2  big-endian, covers all bytes except these last two
    """
    data_bytes = (command + '\n').encode('ascii')
    data_size  = len(data_bytes)

    # Total packet size = header (16) + unique_id (4) + tt_type (1) + data_size_field (2)
    #                   + data (N) + crc (2)
    packet_size = FP_PKT_HEADER_SIZE + 4 + 1 + 2 + data_size + 2

    pkt = bytearray(packet_size)
    i = 0

    # Header
    pkt[i] = FP_SYNCH;                                          i += 1
    pkt[i] = FP_TEXT;                                           i += 1
    struct.pack_into('>H', pkt, i, packet_size);                i += 2
    struct.pack_into('>I', pkt, i, destination_id);             i += 4
    struct.pack_into('>I', pkt, i, FP_THERAVISION_ID);          i += 4
    pkt[i] = 0;                                                 i += 1  # sequence_number
    pkt[i] = 0;                                                 i += 1  # broadcast_destination
    pkt[i] = 0;                                                 i += 1  # broadcast_depth
    pkt[i] = 0;                                                 i += 1  # repeat_count

    # Payload
    struct.pack_into('>I', pkt, i, FP_UNIQUE_ID);               i += 4
    pkt[i] = FP_TT_KEYBOARD;                                    i += 1
    struct.pack_into('>H', pkt, i, data_size);                  i += 2
    pkt[i:i + data_size] = data_bytes;                          i += data_size

    # CRC over all bytes except the last two (the CRC field itself).
    # repeat_count (byte 15) is already 0, matching the CRC calculation rule.
    crc = _crc16_fsd(bytes(pkt[:packet_size - 2]))
    struct.pack_into('>H', pkt, packet_size - 2, crc)

    return bytes(pkt)


# ---------------------------------------------------------------------------
# CSV output (Remote_Terminal calibrate_rfg_N.csv format)
# ---------------------------------------------------------------------------

def write_calibrate_csv(path, cmds):
    """
    Write commands as comma-delimited lines matching the Remote_Terminal format.
    Remote_Terminal replaces all commas with spaces before sending each line,
    so 'CALIBRATE,TABLE,0' becomes 'CALIBRATE TABLE 0' on the wire.
    """
    with open(path, 'w', newline='\r\n') as fh:
        for cmd in cmds:
            fh.write(','.join(cmd.split()) + '\r\n')
    print(f'  Written: {path}')


# ---------------------------------------------------------------------------
# Serial upload via FSD
# ---------------------------------------------------------------------------

def send_commands_fsd(port, rfg_idx, cmds, delay):
    try:
        import serial
    except ImportError:
        print('ERROR: pyserial not installed.  Run: pip install pyserial', file=sys.stderr)
        return

    destination_id = FP_RF_GENERATOR_ID + rfg_idx
    print(f'  Opening {port} at {FP_BAUD} baud '
          f'(destination_id=0x{destination_id:08X} for RFG{rfg_idx}) ...')

    with serial.Serial(port, FP_BAUD, timeout=1) as ser:
        for cmd in cmds:
            pkt = build_fsd_text_packet(destination_id, cmd)
            ser.write(pkt)
            print(f'    >> {cmd}')
            time.sleep(delay)
            # Drain any response bytes; print printable ASCII for diagnostics.
            raw = ser.read_all()
            if raw:
                printable = ''.join(
                    chr(b) if 0x20 <= b < 0x7F or b in (0x0A, 0x0D) else f'<{b:02X}>'
                    for b in raw
                )
                for line in printable.splitlines():
                    line = line.strip()
                    if line:
                        print(f'    << {line}')

    print(f'  Upload to {port} (RFG{rfg_idx}) complete.')


# ---------------------------------------------------------------------------
# Per-RFG processing
# ---------------------------------------------------------------------------

def process_rfg(rfg_idx, directory, port, out_dir, delay):
    print(f'--- RFG{rfg_idx} ---')

    def load(ch_idx, direction):
        prefix = _FILE_PREFIXES[(rfg_idx, ch_idx, direction)]
        path   = find_log(directory, prefix)
        if path is None:
            print(f'  WARNING: no file found for {prefix}* in {directory}')
            return []
        print(f'  Loading {os.path.basename(path)}')
        rows = load_log(path)
        print(f'    {len(rows)} rows')
        return rows

    fwd0 = load(0, 'F')
    ref0 = load(0, 'R')
    fwd1 = load(1, 'F')
    ref1 = load(1, 'R')

    valid_freqs, excluded = valid_frequencies(fwd0 + fwd1)
    for freq, (conv, tot) in sorted(excluded.items()):
        print(f'  SKIP {freq} kHz: only {conv}/{tot} setpoints converged '
              f'(need {MIN_CONVERGED_ROWS}, threshold {CONVERGENCE_THRESHOLD*100:.0f}%) '
              f'-- amplifier limit?')

    tables = build_tables(fwd0, ref0, fwd1, ref1, valid_freqs=valid_freqs)
    if not tables:
        print(f'  ERROR: no frequency data found for RFG{rfg_idx}, skipping.')
        return

    if len(tables) > MAX_CALIBRATION_TABLES:
        print(f'  WARNING: {len(tables)} frequencies exceed firmware maximum '
              f'{MAX_CALIBRATION_TABLES}, truncating.')
        tables = tables[:MAX_CALIBRATION_TABLES]

    print(f'  {len(tables)} table slot(s): '
          + ', '.join(f'{t["freq_khz"]} kHz' for t in tables))

    drive_ch0 = fit_drive_poly(fwd0)
    drive_ch1 = fit_drive_poly(fwd1)

    for label, coeffs in (('A', drive_ch0), ('B', drive_ch1)):
        if coeffs is None:
            print(f'  Drive poly CH_{label}: insufficient data (need >=6 rows with drive_sum > 0)')
        else:
            c  = coeffs[2]
            ok = SENTINEL_C_MIN < c < SENTINEL_C_MAX
            print(f'  Drive poly CH_{label}: c={c:.6f}  sentinel={"PASS" if ok else "FAIL"}')

    cmds = build_commands(tables, drive_ch0, drive_ch1)
    print(f'  {len(cmds)} CALIBRATE commands generated')

    if port is None:
        print()
        print('  [dry run -- no --port specified, printing commands only]')
        for cmd in cmds:
            print(f'    {cmd}')
        print()
    else:
        send_commands_fsd(port, rfg_idx, cmds, delay)

    csv_path = os.path.join(out_dir, f'calibrate_rfg_{rfg_idx}.csv')
    write_calibrate_csv(csv_path, cmds)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def parse_args():
    p = argparse.ArgumentParser(
        description='Upload RFG calibration via FSD serial bus (replaces Remote_Terminal manual steps)'
    )
    p.add_argument('--dir', required=True,
                   help='Directory containing rfg_NxF/rfg_NxR log files')
    p.add_argument('--port', default=None,
                   help='FSD serial port (e.g. COM60, 921600 baud). Omit for dry run.')
    p.add_argument('--rfg', default='0,1',
                   help='Comma-separated list of RFG board numbers to program (default: 0,1)')
    p.add_argument('--out', default=None,
                   help='Output directory for calibrate_rfg_N.csv (default: same as --dir)')
    p.add_argument('--delay', type=float, default=CMD_DELAY_S,
                   help=f'Seconds between FSD packets (default {CMD_DELAY_S})')
    return p.parse_args()


def main():
    args   = parse_args()
    directory = os.path.abspath(args.dir)
    out_dir   = os.path.abspath(args.out) if args.out else directory

    if not os.path.isdir(directory):
        print(f'ERROR: directory not found: {directory}', file=sys.stderr)
        sys.exit(1)

    try:
        rfg_indices = [int(x.strip()) for x in args.rfg.split(',')]
    except ValueError:
        print(f'ERROR: --rfg must be comma-separated integers, got: {args.rfg}', file=sys.stderr)
        sys.exit(1)

    os.makedirs(out_dir, exist_ok=True)

    print(f'Input directory : {directory}')
    print(f'Output directory: {out_dir}')
    if args.port:
        print(f'FSD port        : {args.port} at {FP_BAUD} baud')
    else:
        print('FSD port        : (dry run)')
    print(f'RFG boards      : {rfg_indices}')
    print()

    for rfg_idx in rfg_indices:
        process_rfg(rfg_idx, directory, args.port, out_dir, args.delay)
        print()

    print('Done.')


if __name__ == '__main__':
    main()
