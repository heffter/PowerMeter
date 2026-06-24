#!/usr/bin/env python3
"""
calibrate_tvn.py -- Batch RFG sensor and drive calibration uploader.

Reads 8 auto-calibration log files produced by Remote_Control auto-cal sweep
(one file per RFG/channel/direction), fits sensor polynomials (FORWARD/REFLECTED)
per frequency range and a per-channel drive polynomial (POWER), then uploads all
CALIBRATE commands to RFG0 and RFG1 via their ASCII debug serial ports and writes
calibrate_rfg_0.csv / calibrate_rfg_1.csv as side-effect files for record-keeping.

Replaces the Excel-based workflow (automate_calibration.bat + automate_excel.ps1)
and eliminates the manual Remote Terminal upload step.

Usage:
    python calibrate_tvn.py --dir C:\\cal_data\\TVN-42 --port0 COM5 --port1 COM6
    python calibrate_tvn.py --dir C:\\cal_data\\TVN-42          # dry run (no upload)

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
    where x = sqrt(ams_watts)   -- the raw ADC domain input

Drive polynomial model (bivariate, stored in TABLE 0, per channel):
    volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)
    where f = freq_MHz, P = bird_fwd_watts, volt = sqrt(drive_sum) * DRIVE_VOLT_SCALE

Firmware sentinel: drive polynomial only activates when 0.01 < c < 10.0.

Prerequisites:
    pip install numpy pyserial
"""

import argparse
import glob
import math
import os
import sys
import time

import numpy as np

# Must match firmware rf_generator.hc constants
DRIVE_VOLT_SCALE = 2.5 / math.sqrt(60.0)
SENTINEL_C_MIN = 0.01
SENTINEL_C_MAX = 10.0

BAUD = 256000
CMD_DELAY_S = 0.4          # seconds between serial commands
MAX_CALIBRATION_TABLES = 16

# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

# Maps (rfg_index, channel_index, direction) -> log file name prefix
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
    """Return path to first file matching prefix*.{log,csv} in directory, or None."""
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
        channel       int  rfg_channel (0=A, 1=B)
        freq_khz      int  RF frequency in kHz
        pwr_setpoint  float requested power (W)
        ams_fwd       float RFG internal forward ADC reading (W, identity poly)
        bird_fwd      float reference meter forward power (W)
        ams_ref       float RFG internal reflected ADC reading (W, identity poly)
        bird_ref      float reference meter reflected power (W)
        drive_sum     float DRIVESUM snapshot (0.0 when not available)
    """
    rows = []
    with open(path, newline='', encoding='utf-8', errors='replace') as fh:
        raw_lines = fh.readlines()

    # First line is the file path; second line is the column header.
    # Find the header line by scanning for 'Channel'.
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
# Polynomial fitting
# ---------------------------------------------------------------------------

_IDENTITY_POLY = (0.0, 0.0, 1.0, 0.0, 0.0)


def fit_sensor_poly(ams_values, bird_values):
    """
    Fit 4th-degree polynomial: bird = a + b*x + c*x^2 + d*x^3 + e*x^4
    where x = sqrt(ams) -- the raw ADC input domain.

    The identity polynomial (a=0, b=0, c=1, d=0, e=0) maps x^2 = ams -> ams,
    so uncalibrated output equals the AMS reading. A real fit corrects for
    detector nonlinearity.

    Returns (a, b, c, d, e) or the identity polynomial when data is too sparse.
    """
    valid = [(a, b) for a, b in zip(ams_values, bird_values) if a >= 0 and b > 0]
    if len(valid) < 3:
        return _IDENTITY_POLY

    x = np.array([math.sqrt(a) for a, _ in valid])
    y = np.array([b for _, b in valid])

    # np.polyfit returns descending order: [e, d, c, b, a] for degree 4
    p = np.polyfit(x, y, deg=4)
    return (float(p[4]), float(p[3]), float(p[2]), float(p[1]), float(p[0]))


def fit_drive_poly(rows, min_power=1.0):
    """
    Fit bivariate drive polynomial from a forward-relay measurement set.

    Model: volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)
    where f = freq_MHz, P = bird_fwd_W, volt = sqrt(drive_sum) * DRIVE_VOLT_SCALE

    Returns (a, b, c, d, e) or None when there are fewer than 6 usable points.
    """
    pts = []
    for row in rows:
        if row['drive_sum'] <= 0.0 or row['bird_fwd'] < min_power:
            continue
        f = row['freq_khz'] / 1000.0
        p = row['bird_fwd']
        volt = math.sqrt(row['drive_sum']) * DRIVE_VOLT_SCALE
        pts.append((f, p, volt))

    if len(pts) < 6:
        return None

    A = np.array([[1.0, f, math.sqrt(p), f * f, f * math.sqrt(p)]
                  for f, p, _ in pts])
    y = np.array([v for _, _, v in pts])
    coeffs, _, _, _ = np.linalg.lstsq(A, y, rcond=None)
    return tuple(float(c) for c in coeffs)


# ---------------------------------------------------------------------------
# Frequency table structure
# ---------------------------------------------------------------------------

def build_tables(fwd0_rows, ref0_rows, fwd1_rows, ref1_rows):
    """
    Group measurements by frequency and fit per-slot polynomials.

    One firmware TABLE slot is created per distinct calibration frequency.
    Frequency range boundaries are set at midpoints between adjacent frequencies,
    with half-step margins on the outer edges.

    Returns list of table dicts (sorted ascending by frequency):
        freq_khz  int
        freq_lo   int  RANGE lower bound (kHz)
        freq_hi   int  RANGE upper bound (kHz)
        fwd_max   int  RANGE forward power maximum (W, rounded up)
        ref_max   int  RANGE reflected power maximum (W, rounded up)
        fwd0      tuple  (a,b,c,d,e) forward polynomial for channel A
        ref0      tuple  reflected polynomial for channel A
        fwd1      tuple  forward polynomial for channel B
        ref1      tuple  reflected polynomial for channel B
    """
    all_rows = fwd0_rows + ref0_rows + fwd1_rows + ref1_rows
    freqs = sorted(set(r['freq_khz'] for r in all_rows))
    if not freqs:
        return []

    tables = []
    n = len(freqs)
    for i, freq in enumerate(freqs):
        # Frequency range boundaries
        half_prev = (freq - freqs[i - 1]) // 2 if i > 0 else (freqs[1] - freq) // 2 if n > 1 else 500
        half_next = (freqs[i + 1] - freq) // 2 if i < n - 1 else (freq - freqs[i - 1]) // 2 if n > 1 else 500
        f_lo = max(0, freq - half_prev)
        f_hi = freq + half_next

        def at(rows):
            return [r for r in rows if r['freq_khz'] == freq]

        f0 = at(fwd0_rows)
        r0 = at(ref0_rows)
        f1 = at(fwd1_rows)
        r1 = at(ref1_rows)

        fwd_vals = [r['bird_fwd'] for r in f0 + f1 if r['bird_fwd'] > 0]
        ref_vals = [r['bird_ref'] for r in r0 + r1 if r['bird_ref'] > 0]
        fwd_max = max(1, math.ceil(max(fwd_vals))) if fwd_vals else 35
        ref_max = max(1, math.ceil(max(ref_vals))) if ref_vals else 10

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

def _fmt(v):
    """Format a coefficient with 8 decimal places."""
    return f'{v:.8f}'


def _coeff_str(a, b, c, d, e):
    return f'{_fmt(a)} {_fmt(b)} {_fmt(c)} {_fmt(d)} {_fmt(e)}'


def build_commands(tables, drive_ch0, drive_ch1):
    """
    Generate the complete CALIBRATE command sequence for one RFG.

    Sensor polynomials go into table slots 0..N-1 (one per calibration frequency).
    Drive polynomials are stored in TABLE 0 alongside the slot-0 sensor data.

    Returns list of ASCII command strings ready to send at 256000 baud.
    """
    cmds = []

    for slot, tbl in enumerate(tables):
        cmds.append(f'CALIBRATE TABLE {slot}')
        for ch_cmd in ('1', '2'):
            cmds.append(
                f'CALIBRATE {ch_cmd} RANGE '
                f'{tbl["freq_lo"]} {tbl["freq_hi"]} '
                f'0 {tbl["fwd_max"]} '
                f'0 {tbl["ref_max"]} '
                f'0 {tbl["fwd_max"]}'
            )
        cmds.append(f'CALIBRATE 1 FORWARD {_coeff_str(*tbl["fwd0"])}')
        cmds.append(f'CALIBRATE 2 FORWARD {_coeff_str(*tbl["fwd1"])}')
        cmds.append(f'CALIBRATE 1 REFLECTED {_coeff_str(*tbl["ref0"])}')
        cmds.append(f'CALIBRATE 2 REFLECTED {_coeff_str(*tbl["ref1"])}')

        # Drive polynomial lives in TABLE 0 alongside the first sensor slot.
        if slot == 0:
            if drive_ch0 is not None:
                cmds.append(f'CALIBRATE 1 POWER {_coeff_str(*drive_ch0)}')
            if drive_ch1 is not None:
                cmds.append(f'CALIBRATE 2 POWER {_coeff_str(*drive_ch1)}')

        cmds.append('CALIBRATE WRITE')

    return cmds


# ---------------------------------------------------------------------------
# CSV output (Remote_Terminal calibrate_rfg_N.csv format)
# ---------------------------------------------------------------------------

def write_calibrate_csv(path, cmds):
    """
    Write commands as a comma-delimited file matching the format Remote_Terminal
    expects: each ASCII word in a command becomes a comma-separated field.
    Remote_Terminal replaces all commas with spaces before sending each line.
    """
    with open(path, 'w', newline='\r\n') as fh:
        for cmd in cmds:
            # Convert 'CALIBRATE TABLE 0' -> 'calibrate,TABLE,0'
            tokens = cmd.split()
            fh.write(','.join(tokens) + '\r\n')
    print(f'  Written: {path}')


# ---------------------------------------------------------------------------
# Serial upload
# ---------------------------------------------------------------------------

def send_commands(port, baud, cmds, delay):
    try:
        import serial
    except ImportError:
        print(
            'ERROR: pyserial not installed.  Run: pip install pyserial',
            file=sys.stderr,
        )
        return

    print(f'  Opening {port} at {baud} baud ...')
    with serial.Serial(port, baud, timeout=2) as ser:
        for cmd in cmds:
            ser.write((cmd + '\r\n').encode('ascii'))
            print(f'    >> {cmd}')
            time.sleep(delay)
            resp = ser.read_all().decode('ascii', errors='replace').strip()
            for line in resp.splitlines():
                print(f'    << {line}')
    print(f'  Upload to {port} complete.')


# ---------------------------------------------------------------------------
# Per-RFG processing
# ---------------------------------------------------------------------------

def process_rfg(rfg_idx, directory, port, out_dir, baud, delay):
    print(f'--- RFG{rfg_idx} ---')

    keys = {
        (0, 'F'): _FILE_PREFIXES[(rfg_idx, 0, 'F')],
        (0, 'R'): _FILE_PREFIXES[(rfg_idx, 0, 'R')],
        (1, 'F'): _FILE_PREFIXES[(rfg_idx, 1, 'F')],
        (1, 'R'): _FILE_PREFIXES[(rfg_idx, 1, 'R')],
    }

    def load(key):
        prefix = keys[key]
        path = find_log(directory, prefix)
        if path is None:
            print(f'  WARNING: no file found for {prefix}* in {directory}')
            return []
        print(f'  Loading {os.path.basename(path)}')
        rows = load_log(path)
        print(f'    {len(rows)} rows')
        return rows

    fwd0 = load((0, 'F'))   # channel A forward
    ref0 = load((0, 'R'))   # channel A reflected
    fwd1 = load((1, 'F'))   # channel B forward
    ref1 = load((1, 'R'))   # channel B reflected

    tables = build_tables(fwd0, ref0, fwd1, ref1)
    if not tables:
        print(f'  ERROR: no frequency data found for RFG{rfg_idx}, skipping.')
        return

    if len(tables) > MAX_CALIBRATION_TABLES:
        print(
            f'  WARNING: {len(tables)} frequencies exceed firmware maximum '
            f'{MAX_CALIBRATION_TABLES}, truncating to lowest {MAX_CALIBRATION_TABLES}.'
        )
        tables = tables[:MAX_CALIBRATION_TABLES]

    print(
        f'  {len(tables)} table slot(s): '
        + ', '.join(f'{t["freq_khz"]} kHz' for t in tables)
    )

    drive_ch0 = fit_drive_poly(fwd0)
    drive_ch1 = fit_drive_poly(fwd1)

    for ch_label, drive_coeffs in (('A', drive_ch0), ('B', drive_ch1)):
        if drive_coeffs is None:
            print(f'  Drive poly CH_{ch_label}: insufficient data (need >=6 rows with drive_sum > 0)')
        else:
            c = drive_coeffs[2]
            ok = SENTINEL_C_MIN < c < SENTINEL_C_MAX
            print(f'  Drive poly CH_{ch_label}: c={c:.6f}  sentinel={"PASS" if ok else "FAIL"}')

    cmds = build_commands(tables, drive_ch0, drive_ch1)
    print(f'  {len(cmds)} CALIBRATE commands generated')

    if port is None:
        print()
        print('  [dry run -- no --port specified, printing commands only]')
        for cmd in cmds:
            print(f'    {cmd}')
        print()
    else:
        send_commands(port, baud, cmds, delay)

    csv_path = os.path.join(out_dir, f'calibrate_rfg_{rfg_idx}.csv')
    write_calibrate_csv(csv_path, cmds)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def parse_args():
    p = argparse.ArgumentParser(
        description='Upload RFG sensor + drive calibration from Remote_Control log files'
    )
    p.add_argument(
        '--dir', required=True,
        help='Directory containing rfg_NxF/rfg_NxR log files',
    )
    p.add_argument('--port0', default=None, help='COM port for RFG0 (e.g. COM5)')
    p.add_argument('--port1', default=None, help='COM port for RFG1 (e.g. COM6)')
    p.add_argument(
        '--out', default=None,
        help='Output directory for calibrate_rfg_N.csv (default: same as --dir)',
    )
    p.add_argument('--baud', type=int, default=BAUD, help=f'Serial baud rate (default {BAUD})')
    p.add_argument(
        '--delay', type=float, default=CMD_DELAY_S,
        help=f'Seconds between serial commands (default {CMD_DELAY_S})',
    )
    return p.parse_args()


def main():
    args = parse_args()
    directory = os.path.abspath(args.dir)
    out_dir = os.path.abspath(args.out) if args.out else directory

    if not os.path.isdir(directory):
        print(f'ERROR: directory not found: {directory}', file=sys.stderr)
        sys.exit(1)

    os.makedirs(out_dir, exist_ok=True)

    print(f'Input directory : {directory}')
    print(f'Output directory: {out_dir}')
    print()

    for rfg_idx in (0, 1):
        port = args.port0 if rfg_idx == 0 else args.port1
        process_rfg(rfg_idx, directory, port, out_dir, args.baud, args.delay)
        print()

    print('Done.')


if __name__ == '__main__':
    main()
