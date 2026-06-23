#!/usr/bin/env python3
"""
Pulse-mode drive calibration polynomial fitter.

Reads a CSV produced by the DRIVESUM sweep, fits the 5-coefficient drive
polynomial, reports residuals, and optionally sends the CALIBRATE commands
directly to the RFG ASCII serial port.

A PDF calibration report is written automatically next to the input CSV
(suppress with --no-report).

Input CSV columns (one header row required):
    freq_field,power_setpoint_w,keysight_w,drive_sum

    freq_field        rf_frequency_khz XML field value (100 Hz units)
                      field 5000 = 500 kHz = 0.5 MHz
    power_setpoint_w  requested power (W); used for labelling only
    keysight_w        actual power measured by Keysight N1914A (W)
    drive_sum         value reported by CALIBRATE DRIVESUM command

Polynomial fitted:
    volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)

    f    = freq_field / 10000.0          (MHz)
    P    = keysight_w                    (W, actual measured)
    volt = sqrt(drive_sum) * DRIVE_VOLT_SCALE

Firmware sentinel: polynomial is only activated when 0.01 < c < 10.0.

Usage:
    python fit_drive_cal.py sweep.csv
    python fit_drive_cal.py sweep.csv --port COM5
    python fit_drive_cal.py sweep.csv --port COM5 --channel 1
    python fit_drive_cal.py sweep.csv --plot
    python fit_drive_cal.py sweep.csv --no-report

Prerequisites:
    pip install numpy
    pip install matplotlib   (required for report and --plot)
    pip install pyserial     (optional, for --port)
"""

import argparse
import csv
import datetime
import math
import os
import sys

import numpy as np

# Must match RF_POWER_SCALE_NUMERATOR_VOLTS / sqrt(RF_POWER_SCALE_IMPEDANCE_OHMS)
# in firmware rf_generator.hc and Convert_Calibration.cs.
RF_POWER_SCALE_NUMERATOR_VOLTS = 2.5    # amplifier nominal full-scale voltage (V)
RF_POWER_SCALE_IMPEDANCE_OHMS  = 60.0   # amplifier output impedance (ohms)
DRIVE_VOLT_SCALE = RF_POWER_SCALE_NUMERATOR_VOLTS / math.sqrt(RF_POWER_SCALE_IMPEDANCE_OHMS)

SENTINEL_C_MIN = 0.01
SENTINEL_C_MAX = 10.0


def parse_args():
    p = argparse.ArgumentParser(
        description="Fit pulse-mode drive cal polynomial from DRIVESUM sweep CSV"
    )
    p.add_argument("csv_file", help="Input CSV file from DRIVESUM sweep")
    p.add_argument(
        "--channel",
        type=int,
        choices=[1, 2],
        default=None,
        help="RFG channel to program (1 or 2). Omit to program both channels with the same polynomial.",
    )
    p.add_argument(
        "--port",
        default=None,
        help="Serial port for RFG ASCII interface (e.g. COM5). If omitted, commands are printed only.",
    )
    p.add_argument("--baud", type=int, default=256000)
    p.add_argument(
        "--min-power",
        type=float,
        default=1.0,
        metavar="W",
        help="Discard rows where keysight_w is below this value (default 1.0 W). "
             "Low-power points are noisier and can bias the fit.",
    )
    p.add_argument(
        "--plot",
        action="store_true",
        help="Show interactive residual and fit-quality plots after the report (requires matplotlib).",
    )
    p.add_argument(
        "--no-report",
        action="store_true",
        help="Skip automatic PDF report generation.",
    )
    return p.parse_args()


def load_csv(path, min_power):
    rows = []
    skipped = 0
    with open(path, newline="") as f:
        reader = csv.DictReader(f)
        for lineno, row in enumerate(reader, start=2):
            try:
                freq_field = float(row["freq_field"])
                keysight_w = float(row["keysight_w"])
                drive_sum = float(row["drive_sum"])
            except (KeyError, ValueError) as exc:
                print(f"  [line {lineno}] skipped: {exc}", file=sys.stderr)
                skipped += 1
                continue

            if keysight_w < min_power:
                skipped += 1
                continue
            if drive_sum <= 0.0:
                skipped += 1
                continue

            freq_mhz = freq_field / 10000.0
            volt = math.sqrt(drive_sum) * DRIVE_VOLT_SCALE
            rows.append((freq_mhz, keysight_w, volt))

    print(f"Loaded {len(rows)} rows, skipped {skipped} (below {min_power} W or invalid)")
    return rows


def build_matrix(rows):
    # volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)
    A = []
    y = []
    for freq_mhz, keysight_w, volt in rows:
        f = freq_mhz
        s = math.sqrt(keysight_w)
        A.append([1.0, f, s, f * f, f * s])
        y.append(volt)
    return np.array(A, dtype=float), np.array(y, dtype=float)


def fit(rows):
    A, y = build_matrix(rows)
    coeffs, residuals_ss, rank, sv = np.linalg.lstsq(A, y, rcond=None)
    predicted = A @ coeffs
    residuals = predicted - y
    return coeffs, residuals, rank, sv


def rms(values):
    return math.sqrt(np.mean(np.asarray(values) ** 2))


def print_report(coeffs, residuals, rows):
    a, b, c, d, e = coeffs
    print()
    print("Polynomial fit:  volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)")
    print(f"  a = {a: .8f}")
    print(f"  b = {b: .8f}")
    print(f"  c = {c: .8f}")
    print(f"  d = {d: .8f}")
    print(f"  e = {e: .8f}")
    print()
    print(f"Residuals (volt):  max={np.max(np.abs(residuals)):.4f}  rms={rms(residuals):.4f}")

    mid_volt = np.mean([v for _, _, v in rows])
    if mid_volt > 0:
        frac = rms(residuals) / mid_volt
        print(f"Approx RMS power error (at mid-range drive): ~{frac*100:.1f}%")

    print()
    if SENTINEL_C_MIN < c < SENTINEL_C_MAX:
        print(f"Sentinel check: PASS  (c={c:.6f} is between {SENTINEL_C_MIN} and {SENTINEL_C_MAX})")
        print("Firmware will activate this polynomial on next boot after CALIBRATE WRITE.")
    else:
        print(f"Sentinel check: FAIL  (c={c:.6f} must be between {SENTINEL_C_MIN} and {SENTINEL_C_MAX})")
        print("Firmware will NOT activate this polynomial. Check input data quality.")


def build_commands(coeffs, channels):
    a, b, c, d, e = coeffs
    coeff_str = f"{a:.8f} {b:.8f} {c:.8f} {d:.8f} {e:.8f}"
    cmds = [f"CALIBRATE {ch} POWER {coeff_str}" for ch in channels]
    cmds.append("CALIBRATE WRITE")
    return cmds


def send_commands(port, baud, commands):
    try:
        import serial
        import time
    except ImportError:
        print(
            "\nERROR: pyserial not installed.  Install with:  pip install pyserial\n"
            "Send the commands printed above manually via a terminal.",
            file=sys.stderr,
        )
        return

    print(f"\nOpening {port} at {baud} baud ...")
    with serial.Serial(port, baud, timeout=2) as ser:
        for cmd in commands:
            ser.write((cmd + "\r\n").encode("ascii"))
            print(f"  >> {cmd}")
            import time
            time.sleep(0.4)
            response = ser.read_all().decode("ascii", errors="replace").strip()
            if response:
                for line in response.splitlines():
                    print(f"  << {line}")
    print(
        "\nDone. Reboot the RFG and verify the boot log shows:\n"
        "  pulse_drive_cal: cookie 0x44525643 valid 1"
    )


# ---------------------------------------------------------------------------
# PDF report generation
# ---------------------------------------------------------------------------

def _per_freq_stats(rows, residuals):
    """Return ordered list of (freq_mhz, n, rms_v, max_v, bias_v) per frequency."""
    per = {}
    for i, row in enumerate(rows):
        per.setdefault(row[0], []).append(float(residuals[i]))
    result = []
    for f in sorted(per):
        vals = np.array(per[f])
        result.append((f, len(vals), rms(vals), float(np.max(np.abs(vals))), float(np.mean(vals))))
    return result


def generate_report(csv_path, min_power, rows, coeffs, residuals, channels):
    """Write a 4-page PDF calibration report next to the input CSV."""
    try:
        import matplotlib
        import matplotlib.pyplot as plt
        import matplotlib.gridspec as gridspec
        from matplotlib.backends.backend_pdf import PdfPages
    except ImportError:
        print(
            "matplotlib not installed -- skipping report.  Install with: pip install matplotlib",
            file=sys.stderr,
        )
        return None

    base = os.path.splitext(os.path.abspath(csv_path))[0]
    report_path = base + "_report.pdf"

    a, b, c, d, e = coeffs
    sentinel_ok = SENTINEL_C_MIN < c < SENTINEL_C_MAX
    sentinel_label = "PASS" if sentinel_ok else "FAIL"
    sentinel_color = "#2e7d32" if sentinel_ok else "#c62828"

    A_mat, y_vec = build_matrix(rows)
    predicted = A_mat @ coeffs
    # residuals passed in already = predicted - y, recompute to keep consistent
    res = predicted - y_vec

    rms_v   = rms(res)
    max_v   = float(np.max(np.abs(res)))
    bias_v  = float(np.mean(res))
    mid_volt = float(np.mean(y_vec))
    power_err_pct = (rms_v / mid_volt * 100.0) if mid_volt > 0 else 0.0

    freqs_sorted = sorted(set(r[0] for r in rows))
    n_freqs = len(freqs_sorted)
    freq_idx = {f: i for i, f in enumerate(freqs_sorted)}
    cmap = plt.cm.viridis
    point_colors = [cmap(freq_idx[r[0]] / max(n_freqs - 1, 1)) for r in rows]

    freq_stats = _per_freq_stats(rows, res)
    coeff_str = f"{a:.8f} {b:.8f} {c:.8f} {d:.8f} {e:.8f}"
    cmd_lines = [f"CALIBRATE {ch} POWER {coeff_str}" for ch in channels]
    cmd_lines.append("CALIBRATE WRITE")
    now_str = datetime.datetime.now().strftime("%Y-%m-%d  %H:%M:%S")

    # Scalar mappable for shared colorbars (frequency axis)
    sm = plt.cm.ScalarMappable(
        cmap=cmap,
        norm=matplotlib.colors.Normalize(vmin=freqs_sorted[0], vmax=freqs_sorted[-1]),
    )
    sm.set_array([])

    with PdfPages(report_path) as pdf:

        # ------------------------------------------------------------------ #
        # Page 1 -- Calibration summary                                       #
        # ------------------------------------------------------------------ #
        fig, ax = plt.subplots(figsize=(11, 8.5))
        ax.axis("off")
        fig.patch.set_facecolor("#f7f9fc")

        # Header band
        fig.add_axes([0, 0.88, 1, 0.12]).set_axis_off()
        fig.axes[-1].set_facecolor("#1a237e")
        fig.text(0.5, 0.935, "Pulse-Mode Drive Calibration Report",
                 ha="center", va="center", fontsize=18, fontweight="bold",
                 color="white", transform=fig.transFigure)

        # Body text -- left column
        body = (
            f"Generated   {now_str}\n"
            f"Input file  {os.path.basename(csv_path)}\n"
            f"Points used {len(rows)}   (min_power filter: {min_power} W)\n"
            f"Frequencies {n_freqs}   ({freqs_sorted[0]:.3f} - {freqs_sorted[-1]:.3f} MHz)\n"
            f"Power range {min(r[1] for r in rows):.0f} - {max(r[1] for r in rows):.0f} W"
        )
        ax.text(0.04, 0.83, body, transform=ax.transAxes,
                fontfamily="monospace", fontsize=10, va="top", color="#1a237e",
                bbox=dict(boxstyle="round,pad=0.5", facecolor="white", edgecolor="#1a237e", linewidth=0.8))

        # Sentinel badge
        ax.text(0.76, 0.865, f"Sentinel\n{sentinel_label}",
                transform=ax.transAxes, fontsize=14, fontweight="bold",
                ha="center", va="center", color="white",
                bbox=dict(boxstyle="round,pad=0.6", facecolor=sentinel_color, linewidth=0))

        # Polynomial coefficients
        coeff_text = (
            "POLYNOMIAL\n"
            "  volt = a + b·f + c·√P + d·f² + e·f·√P\n"
            f"  a = {a: .8f}\n"
            f"  b = {b: .8f}\n"
            f"  c = {c: .8f}    ← sentinel ({SENTINEL_C_MIN} < c < {SENTINEL_C_MAX}): {sentinel_label}\n"
            f"  d = {d: .8f}\n"
            f"  e = {e: .8f}"
        )
        ax.text(0.04, 0.62, coeff_text, transform=ax.transAxes,
                fontfamily="monospace", fontsize=10, va="top",
                bbox=dict(boxstyle="round,pad=0.5", facecolor="white", edgecolor="#37474f", linewidth=0.8))

        # Residual summary
        res_color = "#fff3e0" if rms_v > 0.05 else "#e8f5e9"
        res_border = "#e65100" if rms_v > 0.05 else "#2e7d32"
        res_text = (
            "RESIDUALS\n"
            f"  RMS  = {rms_v:.4f} V\n"
            f"  Max  = {max_v:.4f} V\n"
            f"  Bias = {bias_v:+.4f} V\n"
            f"  Approx RMS power error at mid-range: ~{power_err_pct:.1f}%"
        )
        ax.text(0.55, 0.62, res_text, transform=ax.transAxes,
                fontfamily="monospace", fontsize=10, va="top",
                bbox=dict(boxstyle="round,pad=0.5", facecolor=res_color, edgecolor=res_border, linewidth=0.8))

        # Commands block
        cmd_block = "COMMANDS  (RFG ASCII port, 256000 baud)\n" + \
                    "\n".join(f"  {c}" for c in cmd_lines)
        ax.text(0.04, 0.34, cmd_block, transform=ax.transAxes,
                fontfamily="monospace", fontsize=9.5, va="top",
                bbox=dict(boxstyle="round,pad=0.5", facecolor="#e3f2fd", edgecolor="#0d47a1", linewidth=0.8))

        # Threshold guide
        guide = (
            "ACCEPTANCE THRESHOLDS\n"
            "  RMS residual < 0.05 V  -- target (approx 3-5% power error at mid-range)\n"
            "  RMS residual < 0.10 V  -- acceptable\n"
            "  RMS residual > 0.10 V  -- re-run sweep; check for outlier rows (try --min-power 3)"
        )
        ax.text(0.04, 0.14, guide, transform=ax.transAxes,
                fontfamily="monospace", fontsize=9, va="top",
                bbox=dict(boxstyle="round,pad=0.5", facecolor="#fce4ec", edgecolor="#880e4f", linewidth=0.8))

        pdf.savefig(fig, bbox_inches="tight", facecolor=fig.get_facecolor())
        plt.close(fig)

        # ------------------------------------------------------------------ #
        # Page 2 -- Fit quality                                               #
        # ------------------------------------------------------------------ #
        fig, axes = plt.subplots(1, 2, figsize=(14, 6))
        fig.suptitle("Fit Quality", fontsize=14, fontweight="bold")

        volt_measured = y_vec.tolist()
        volt_predicted = predicted.tolist()
        powers_all = [r[1] for r in rows]

        # Left: measured vs predicted
        ax = axes[0]
        ax.scatter(volt_measured, volt_predicted, c=point_colors, s=22, alpha=0.85, linewidths=0)
        vmin = min(volt_measured + volt_predicted)
        vmax = max(volt_measured + volt_predicted)
        ax.plot([vmin, vmax], [vmin, vmax], "r--", linewidth=1.2, label="ideal (y = x)")
        ax.set_xlabel("Measured voltage (V)")
        ax.set_ylabel("Predicted voltage (V)")
        ax.set_title("Measured vs Predicted")
        ax.legend(fontsize=9)
        cbar = fig.colorbar(sm, ax=ax, shrink=0.85, pad=0.02)
        cbar.set_label("Frequency (MHz)", fontsize=9)

        # Right: residuals vs power
        ax = axes[1]
        ax.scatter(powers_all, res, c=point_colors, s=22, alpha=0.85, linewidths=0)
        ax.axhline(0, color="r", linewidth=1.2, linestyle="--")
        ax.axhline( 0.05, color="orange", linewidth=0.8, linestyle=":", label="+0.05 V target")
        ax.axhline(-0.05, color="orange", linewidth=0.8, linestyle=":")
        ax.axhline( 0.10, color="red",    linewidth=0.8, linestyle=":", label="+0.10 V limit")
        ax.axhline(-0.10, color="red",    linewidth=0.8, linestyle=":")
        ax.set_xlabel("Keysight power (W)")
        ax.set_ylabel("Residual (V)   [predicted − measured]")
        ax.set_title("Residuals vs Power")
        ax.legend(fontsize=8)
        cbar2 = fig.colorbar(sm, ax=ax, shrink=0.85, pad=0.02)
        cbar2.set_label("Frequency (MHz)", fontsize=9)

        plt.tight_layout()
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ------------------------------------------------------------------ #
        # Page 3 -- Per-frequency analysis                                    #
        # ------------------------------------------------------------------ #
        fig = plt.figure(figsize=(14, 9))
        fig.suptitle("Per-Frequency Analysis", fontsize=14, fontweight="bold")
        gs = gridspec.GridSpec(2, 1, figure=fig, height_ratios=[1.4, 1], hspace=0.45)

        # Bar chart
        ax_bar = fig.add_subplot(gs[0])
        f_labels = [f"{s[0]:.3f}" for s in freq_stats]
        f_rms    = [s[2] for s in freq_stats]
        f_max    = [s[3] for s in freq_stats]
        bar_colors = [cmap(i / max(n_freqs - 1, 1)) for i in range(n_freqs)]
        x = np.arange(n_freqs)
        w = 0.4
        ax_bar.bar(x - w/2, f_rms, width=w, color=bar_colors, edgecolor="black",
                   linewidth=0.5, label="RMS residual")
        ax_bar.bar(x + w/2, f_max, width=w, color=bar_colors, edgecolor="black",
                   linewidth=0.5, alpha=0.45, label="Max |residual|")
        ax_bar.axhline(0.05, color="orange", linewidth=1.2, linestyle="--", label="0.05 V target")
        ax_bar.axhline(0.10, color="red",    linewidth=1.2, linestyle="--", label="0.10 V limit")
        ax_bar.set_xticks(x)
        ax_bar.set_xticklabels(f_labels, rotation=45, ha="right", fontsize=8)
        ax_bar.set_xlabel("Frequency (MHz)")
        ax_bar.set_ylabel("Residual (V)")
        ax_bar.set_title("RMS and Peak Residual by Frequency")
        ax_bar.legend(fontsize=8)

        # Table
        ax_tbl = fig.add_subplot(gs[1])
        ax_tbl.axis("off")
        col_labels = ["Freq (MHz)", "N pts", "RMS (V)", "Max |err| (V)", "Bias (V)", "Status"]
        table_data = []
        for s in freq_stats:
            status = "OK" if s[2] < 0.05 else ("WARN" if s[2] < 0.10 else "FAIL")
            table_data.append([
                f"{s[0]:.3f}", str(s[1]),
                f"{s[2]:.4f}", f"{s[3]:.4f}", f"{s[4]:+.4f}", status,
            ])
        tbl = ax_tbl.table(cellText=table_data, colLabels=col_labels,
                           loc="center", cellLoc="center")
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(8)
        tbl.scale(1, 1.35)
        # Color header
        for j in range(len(col_labels)):
            tbl[0, j].set_facecolor("#1a237e")
            tbl[0, j].set_text_props(color="white", fontweight="bold")
        # Color data rows by status
        for i, s in enumerate(freq_stats):
            if s[2] > 0.10:
                row_color = "#ffcdd2"
            elif s[2] > 0.05:
                row_color = "#fff9c4"
            else:
                row_color = "#c8e6c9"
            for j in range(len(col_labels)):
                tbl[i + 1, j].set_facecolor(row_color)

        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ------------------------------------------------------------------ #
        # Page 4 -- Calibration surface                                       #
        # ------------------------------------------------------------------ #
        fig, axes = plt.subplots(1, 2, figsize=(14, 6))
        fig.suptitle("Calibration Surface  (fitted polynomial over operating envelope)",
                     fontsize=13, fontweight="bold")

        f_lo, f_hi = freqs_sorted[0], freqs_sorted[-1]
        p_lo = max(0.5, min(r[1] for r in rows) * 0.8)
        p_hi = max(r[1] for r in rows) * 1.05

        f_grid = np.linspace(f_lo, f_hi, 100)
        p_grid = np.linspace(p_lo, p_hi, 80)
        F, P = np.meshgrid(f_grid, p_grid)
        V_poly = a + b * F + c * np.sqrt(P) + d * F**2 + e * F * np.sqrt(P)

        # Left: drive voltage surface
        ax = axes[0]
        cf = ax.contourf(F, P, V_poly, levels=25, cmap="viridis")
        cs = ax.contour(F, P, V_poly, levels=10, colors="white", linewidths=0.4, alpha=0.5)
        ax.clabel(cs, inline=True, fontsize=6, fmt="%.2f V")
        ax.scatter([r[0] for r in rows], [r[1] for r in rows],
                   c="white", s=12, alpha=0.55, linewidths=0, label="meas. points")
        cbar = fig.colorbar(cf, ax=ax, pad=0.02)
        cbar.set_label("Drive voltage (V)", fontsize=9)
        ax.set_xlabel("Frequency (MHz)")
        ax.set_ylabel("Power setpoint (W)")
        ax.set_title("Drive Voltage  volt(f, P)")
        ax.legend(fontsize=8, loc="upper left")

        # Right: calibration benefit -- deviation of polynomial drive from
        # the uncorrected sqrt-only law (DRIVE_VOLT_SCALE * sqrt(P)).
        # Shows WHERE the amplifier nonlinearity is largest and the polynomial
        # correction matters most.
        V_uncal = DRIVE_VOLT_SCALE * np.sqrt(P)
        power_deviation_pct = ((V_poly / V_uncal) ** 2 - 1.0) * 100.0

        ax = axes[1]
        abs_lim = max(5.0, float(np.percentile(np.abs(power_deviation_pct), 98)))
        cf2 = ax.contourf(F, P, power_deviation_pct, levels=25, cmap="RdBu_r",
                          vmin=-abs_lim, vmax=abs_lim)
        ax.contour(F, P, power_deviation_pct, levels=[0], colors="black", linewidths=1.0)
        ax.scatter([r[0] for r in rows], [r[1] for r in rows],
                   c="black", s=12, alpha=0.4, linewidths=0, label="meas. points")
        cbar2 = fig.colorbar(cf2, ax=ax, pad=0.02)
        cbar2.set_label("Power correction vs sqrt-only drive (%)", fontsize=8)
        ax.set_xlabel("Frequency (MHz)")
        ax.set_ylabel("Power setpoint (W)")
        ax.set_title("Calibration Benefit  (red = polynomial drives harder,\nblue = softer than sqrt-only law)")
        ax.legend(fontsize=8, loc="upper left")

        plt.tight_layout()
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

    print(f"\nReport saved: {report_path}")
    return report_path


# ---------------------------------------------------------------------------
# Interactive plots (--plot)
# ---------------------------------------------------------------------------

def plot_results(coeffs, residuals, rows):
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        print("matplotlib not installed; skipping plots.  Install with: pip install matplotlib")
        return

    a, b, c, d, e = coeffs
    freqs = sorted(set(r[0] for r in rows))

    fig, axes = plt.subplots(1, 2, figsize=(12, 5))

    # Left: measured vs predicted volt, coloured by frequency
    ax = axes[0]
    cmap = plt.cm.viridis
    freq_idx = {f: i for i, f in enumerate(freqs)}
    colors = [cmap(freq_idx[r[0]] / max(len(freqs) - 1, 1)) for r in rows]
    volt_measured = [r[2] for r in rows]
    A, y = build_matrix(rows)
    volt_predicted = (A @ coeffs).tolist()
    ax.scatter(volt_measured, volt_predicted, c=colors, s=20, alpha=0.7)
    vmin = min(volt_measured + volt_predicted)
    vmax = max(volt_measured + volt_predicted)
    ax.plot([vmin, vmax], [vmin, vmax], "r--", linewidth=1, label="ideal")
    ax.set_xlabel("Measured volt")
    ax.set_ylabel("Predicted volt")
    ax.set_title("Measured vs Predicted (colour = frequency)")
    ax.legend()

    # Right: residuals vs power, coloured by frequency
    ax = axes[1]
    powers = [r[1] for r in rows]
    ax.scatter(powers, residuals, c=colors, s=20, alpha=0.7)
    ax.axhline(0, color="r", linewidth=1)
    ax.set_xlabel("Keysight power (W)")
    ax.set_ylabel("Residual (volt)")
    ax.set_title("Residuals vs Power (colour = frequency)")

    plt.tight_layout()
    plt.show()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    args = parse_args()

    rows = load_csv(args.csv_file, args.min_power)
    if len(rows) < 6:
        print(f"ERROR: {len(rows)} usable rows -- need at least 6 to fit 5 coefficients.")
        sys.exit(1)

    coeffs, residuals, rank, sv = fit(rows)

    if rank < 5:
        print(
            f"WARNING: matrix rank {rank} < 5. The design matrix is under-determined.\n"
            "         Likely cause: not enough distinct power levels in the sweep.\n"
            "         Add more power setpoints so sqrt(P) varies across rows."
        )

    print_report(coeffs, residuals, rows)

    channels = [1, 2] if args.channel is None else [args.channel]
    commands = build_commands(coeffs, channels)

    print("--- Commands for RFG ASCII port (256000 baud) ---")
    for cmd in commands:
        print(cmd)
    print("-------------------------------------------------")

    if not args.no_report:
        generate_report(args.csv_file, args.min_power, rows, coeffs, residuals, channels)

    if args.port:
        send_commands(args.port, args.baud, commands)

    if args.plot:
        plot_results(coeffs, residuals, rows)


if __name__ == "__main__":
    main()
