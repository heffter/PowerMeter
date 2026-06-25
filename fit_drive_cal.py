#!/usr/bin/env python3
"""
Pulse-mode drive calibration polynomial fitter.

Reads a CSV produced by the DRIVESUM sweep, fits the 5-coefficient drive
polynomial, reports residuals, and optionally sends the CALIBRATE commands
directly to the RFG ASCII serial port.

A PDF calibration report is written automatically next to the input CSV
(suppress with --no-report).  Use --png to also save per-page PNGs for
review without a PDF renderer.

Input CSV formats (auto-detected from header row):

    Native format:
        freq_field,power_setpoint_w,keysight_w,drive_sum

        freq_field        rf_frequency_khz XML field value (100 Hz units);
                          field 5000 = 500 kHz = 0.5 MHz
        power_setpoint_w  requested power (W); used for labelling only
        keysight_w        actual power measured by Keysight N1914A (W)
        drive_sum         value reported by CALIBRATE DRIVESUM command

    Remote_Control format (auto-detected when header contains 'Frequency' and 'Drive-Sum'):
        Channel,Frequency,PWR-Control-Level,AMS-FWD-PWR,Bird-FWD-PWR,...,Drive-Sum

        Channel           RFG channel, 0-indexed in the log (CH_A=0, CH_B=1).
                          --channel is 1-indexed (1=CH_A, 2=CH_B).
        Frequency         RF frequency in kHz (e.g. 5000 for 500 kHz)
        Bird-FWD-PWR      Keysight N1914A ch1 forward power reading (W) -- Keysight
                          readings flow into the Bird reference columns via R3 mapping;
                          if a 'Keysight - FWD - PWR' column is present it is preferred
        Drive-Sum         value reported by CALIBRATE DRIVESUM command

Polynomial fitted:
    volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)

    f    = freq_field / 10000.0   (MHz)  [native]
         = frequency_khz / 1000.0 (MHz)  [Remote_Control]
    P    = keysight_w                    (W, actual measured)
    volt = sqrt(drive_sum) * DRIVE_VOLT_SCALE

Firmware sentinel: polynomial is only activated when 0.01 < c < 10.0.

Usage:
    python fit_drive_cal.py sweep.csv
    python fit_drive_cal.py sweep.csv --port COM5
    python fit_drive_cal.py sweep.csv --port COM5 --channel 1
    python fit_drive_cal.py sweep.csv --plot
    python fit_drive_cal.py sweep.csv --png
    python fit_drive_cal.py sweep.csv --no-report

Prerequisites:
    pip install numpy
    pip install matplotlib   (required for report, --plot, and --png)
    pip install pyserial     (optional, for --port)
"""

import argparse
import csv
import datetime
import io
import math
import os
import re
import sys

import numpy as np

# Must match RF_POWER_SCALE_NUMERATOR_VOLTS / sqrt(RF_POWER_SCALE_IMPEDANCE_OHMS)
# in firmware rf_generator.hc and Convert_Calibration.cs.
RF_POWER_SCALE_NUMERATOR_VOLTS = 2.5    # amplifier nominal full-scale voltage (V)
RF_POWER_SCALE_IMPEDANCE_OHMS  = 60.0   # amplifier output impedance (ohms)
DRIVE_VOLT_SCALE = RF_POWER_SCALE_NUMERATOR_VOLTS / math.sqrt(RF_POWER_SCALE_IMPEDANCE_OHMS)

SENTINEL_C_MIN = 0.01
SENTINEL_C_MAX = 10.0

# Acceptance thresholds for residuals (in drive voltage, V)
THRESHOLD_TARGET_V = 0.05   # aim for below this (~3-5% power error)
THRESHOLD_LIMIT_V  = 0.10   # re-sweep recommended above this


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
        help="RFG channel to program (1=CH_A, 2=CH_B). Omit to use all rows and "
             "program both channels with the same polynomial.",
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
    p.add_argument(
        "--png",
        action="store_true",
        help="Also save each report page as a PNG next to the PDF. "
             "Useful for review without a PDF renderer installed.",
    )
    return p.parse_args()


def _detect_csv_format(headers):
    """Return (freq_col, power_col, drive_col, freq_scale, channel_col, fmt_name)."""
    if "freq_field" in headers:
        # Native format: field value in 100 Hz units -> MHz = field / 10000
        return ("freq_field", "keysight_w", "drive_sum", 1.0 / 10000.0, None, "native")

    if "Frequency" in headers and "Drive-Sum" in headers:
        # Remote_Control CSV (kHz -> MHz = kHz / 1000).
        # Prefer the explicit Keysight column when present; otherwise fall back to
        # Bird-FWD-PWR which is where R3 maps Keysight ch1.
        power_col = (
            "Keysight-FWD-PWR" if "Keysight-FWD-PWR" in headers
            else "Bird-FWD-PWR"
        )
        channel_col = "Channel" if "Channel" in headers else None
        return ("Frequency", power_col, "Drive-Sum", 1.0 / 1000.0, channel_col, "Remote_Control")

    print(
        "ERROR: unrecognised CSV format.\n"
        f"  Headers found: {headers}\n"
        "  Expected 'freq_field' (native) or 'Frequency'+'Drive-Sum' (Remote_Control).",
        file=sys.stderr,
    )
    sys.exit(1)


def _normalize_header(name):
    """Strip spaces around hyphens so 'Drive - Sum' matches 'Drive-Sum'."""
    return re.sub(r'\s*-\s*', '-', name).strip()


def load_csv(path, min_power, channel_filter=None):
    rows = []
    skipped = 0
    with open(path, newline="", encoding="utf-8", errors="replace") as f:
        raw_lines = f.readlines()

    # Skip any preamble lines (e.g. the file-path header Remote_Control prepends)
    # and find the real CSV header: the first line that names our known columns.
    header_idx = 0
    for i, line in enumerate(raw_lines):
        if 'Channel' in line and 'Frequency' in line:
            header_idx = i
            break
        if 'freq_field' in line:
            header_idx = i
            break

    csv_text = "".join(raw_lines[header_idx:])
    with io.StringIO(csv_text) as buf:
        reader = csv.DictReader(buf)
        # Normalize header names: strip spaces around hyphens.
        raw_fieldnames = list(reader.fieldnames or [])
        norm_map = {_normalize_header(h): h for h in raw_fieldnames}
        headers = list(norm_map.keys())
        freq_col, power_col, drive_col, freq_scale, channel_col, fmt = \
            _detect_csv_format(headers)

        # Map normalized column names back to raw DictReader field names.
        def col(name):
            return norm_map.get(name, name)

        print(f"CSV format: {fmt}  (power column: '{power_col}')")
        if channel_col and channel_filter is not None:
            # Remote_Control log uses 0-indexed channels (CH_A=0, CH_B=1).
            # --channel is 1-indexed (1=CH_A, 2=CH_B), matching firmware numbering.
            print(f"Filtering to channel {channel_filter} (log value {channel_filter - 1})")

        for lineno, row in enumerate(reader, start=2):
            try:
                if channel_col and channel_filter is not None:
                    # Subtract 1 to convert 1-indexed --channel to 0-indexed log value.
                    log_ch = str(channel_filter - 1)
                    if row.get(col(channel_col), "").strip() != log_ch:
                        skipped += 1
                        continue

                freq_raw   = float(row[col(freq_col)])
                keysight_w = float(row[col(power_col)])
                drive_sum  = float(row[col(drive_col)])
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

            freq_mhz = freq_raw * freq_scale
            volt = math.sqrt(drive_sum) * DRIVE_VOLT_SCALE
            rows.append((freq_mhz, keysight_w, volt))

    print(f"Loaded {len(rows)} rows, skipped {skipped} (below {min_power} W or invalid)")
    return rows


def build_matrix(rows):
    """Build the least-squares design matrix A and response vector y.

    Model: volt = a + b*f + c*sqrt(P) + d*f^2 + e*f*sqrt(P)
    Columns of A: [1, f, sqrt(P), f^2, f*sqrt(P)]
    """
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


def compute_fit_statistics(rows, coeffs, residuals):
    """Compute regression statistics for the fitted polynomial.

    Returns a dict with headline and diagnostic statistics suitable for a
    calibration report.  All quantities are computed from the OLS design matrix
    so they are internally consistent with the fitted coefficients.

    Keys: n, p, dof, SSE, SST, SSR, s, RMSE, R2, R2_adj, max_resid, bias,
          F_stat, t_crit, coeff_SE, coeff_CI_lo, coeff_CI_hi, coeff_t,
          cond_number, max_vif, vifs, power_err_rms_pct, power_err_max_pct,
          U_k2, mid_volt.
    """
    A, y = build_matrix(rows)
    n = len(y)
    p = A.shape[1]   # 5 coefficients
    dof = n - p

    predicted = A @ coeffs
    res = predicted - y

    SSE = float(res @ res)
    SST = float(np.sum((y - np.mean(y)) ** 2))
    SSR = SST - SSE

    # Standard error of estimate (residual standard deviation): the headline
    # goodness-of-fit number, in the same units as the response (volts).
    s = math.sqrt(SSE / dof) if dof > 0 else float("nan")
    RMSE = math.sqrt(SSE / n)  # denominator n, not dof -- reported with that label

    R2 = 1.0 - SSE / SST if SST > 0 else float("nan")
    R2_adj = (
        1.0 - (SSE / dof) / (SST / (n - 1))
        if (dof > 0 and SST > 0)
        else float("nan")
    )

    # Per-coefficient uncertainty via the OLS covariance: Cov(b_hat) = s^2 (A'A)^-1.
    try:
        AtA_inv = np.linalg.inv(A.T @ A)
        cov = (s ** 2) * AtA_inv
        coeff_SE = np.sqrt(np.diag(cov))
    except np.linalg.LinAlgError:
        coeff_SE = np.full(p, float("nan"))

    # t critical value for 95% CI (two-tailed).
    # scipy.stats.t is preferred; fall back to an accurate approximation otherwise.
    try:
        from scipy.stats import t as scipy_t
        t_crit = float(scipy_t.ppf(0.975, dof))
    except ImportError:
        # For dof >= 60 the t-distribution is within ~0.5% of the normal 1.96;
        # 1.984 matches NIST tables for dof = 100.
        t_crit = 1.984 if dof >= 60 else 2.0

    coeff_CI_lo = coeffs - t_crit * coeff_SE
    coeff_CI_hi = coeffs + t_crit * coeff_SE
    coeff_t = np.where(coeff_SE != 0, coeffs / coeff_SE, float("nan"))

    # Overall F-statistic (model vs intercept-only null hypothesis).
    k = p - 1   # non-intercept terms
    F_stat = (SSR / k) / (SSE / dof) if (dof > 0 and SSE > 0) else float("nan")

    # Design-matrix condition number: normalize columns to unit std first so the
    # intercept column (all ones, zero variance) does not dominate.
    col_scales = np.std(A, axis=0)
    col_scales[col_scales == 0] = 1.0
    cond_number = float(np.linalg.cond(A / col_scales))

    # Variance Inflation Factors for the non-intercept predictors.
    # VIF_j = diag(inv(R)) where R is the correlation matrix of the non-constant
    # columns.  VIF > 10 signals harmful multicollinearity (NIST / Kutner ch. 10).
    X = A[:, 1:]   # drop intercept column
    try:
        corr = np.corrcoef(X.T)
        vifs = np.diag(np.linalg.inv(corr))
        max_vif = float(np.max(vifs))
    except np.linalg.LinAlgError:
        vifs = np.full(X.shape[1], float("nan"))
        max_vif = float("nan")

    # Application-level error in % power.
    # For small perturbations: dP/P ~= 2*dV/V  (P proportional to V^2
    # through the fixed-impedance output model), referenced to mean drive volt.
    mid_volt = float(np.mean(y))
    if mid_volt > 0:
        power_err_rms_pct = (rms(res) / mid_volt) * 100.0
        power_err_max_pct = (float(np.max(np.abs(res))) / mid_volt) * 100.0
    else:
        power_err_rms_pct = float("nan")
        power_err_max_pct = float("nan")

    # GUM/JCGM 100 expanded uncertainty: U = k * u_c, k=2, coverage ~95%.
    # u_c here is s (Type-A lack-of-fit).  A full budget adds Type-B contributions
    # from the Keysight N1914A and the signal generator in quadrature.
    U_k2 = 2.0 * s

    return dict(
        n=n, p=p, dof=dof,
        SSE=SSE, SST=SST, SSR=SSR,
        s=s, RMSE=RMSE,
        R2=R2, R2_adj=R2_adj,
        max_resid=float(np.max(np.abs(res))),
        bias=float(np.mean(res)),
        F_stat=F_stat,
        t_crit=t_crit,
        coeff_SE=coeff_SE,
        coeff_CI_lo=coeff_CI_lo,
        coeff_CI_hi=coeff_CI_hi,
        coeff_t=coeff_t,
        cond_number=cond_number,
        max_vif=max_vif,
        vifs=vifs,
        power_err_rms_pct=power_err_rms_pct,
        power_err_max_pct=power_err_max_pct,
        U_k2=U_k2,
        mid_volt=mid_volt,
    )


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
    cmds = [
        "CALIBRATE READ",       # reload flash -> RAM so FORWARD/REFLECTED constants are preserved
        "CALIBRATE TABLE 0",
    ]
    cmds += [f"CALIBRATE {ch} POWER {coeff_str}" for ch in channels]
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
# PDF report: shared style constants
# ---------------------------------------------------------------------------

# Neutral, restrained professional palette (Bootstrap-derived; matches the
# operator PowerMeter.py GUI so reports and UI share a coherent look).
_C_TEXT   = "#343a40"   # body text / axis labels
_C_MUTED  = "#6c757d"   # secondary labels, ticks, footer / header text
_C_GRID   = "#d0d4d9"   # gridlines and separator rules
_C_ACCENT = "#2c7be5"   # single blue accent for data points and headings
_C_PASS   = "#28a745"   # PASS / OK indicator (text only; never a filled box)
_C_FAIL   = "#dc3545"   # FAIL indicator
_C_WARN   = "#e67e22"   # WARN / marginal indicator

# rcParams applied for every report figure via matplotlib.rc_context().
# Setting pdf.fonttype=42 embeds TrueType glyphs so the PDF text is
# selectable and renders identically on all viewers (Type-3 default can
# mis-render on some printers).
_REPORT_RC = {
    "font.family": "sans-serif",
    "font.sans-serif": ["DejaVu Sans", "Helvetica", "Arial"],
    "font.size": 9,
    "axes.titlesize": 10,
    "axes.titleweight": "bold",
    "axes.labelsize": 9,
    "xtick.labelsize": 8,
    "ytick.labelsize": 8,
    "legend.fontsize": 8,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.linewidth": 0.7,
    "axes.edgecolor": "#999999",
    "axes.labelcolor": _C_TEXT,
    "text.color": _C_TEXT,
    "axes.grid": True,
    "axes.grid.axis": "y",
    "grid.color": _C_GRID,
    "grid.linewidth": 0.5,
    "grid.alpha": 0.9,
    "xtick.direction": "out",
    "ytick.direction": "out",
    "xtick.major.size": 3,
    "ytick.major.size": 3,
    "xtick.color": _C_MUTED,
    "ytick.color": _C_MUTED,
    "lines.linewidth": 1.4,
    "legend.frameon": False,
    "figure.dpi": 110,
    "savefig.dpi": 200,
    "pdf.fonttype": 42,
    "ps.fonttype": 42,
}


# ---------------------------------------------------------------------------
# PDF report: helper functions
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


def _normal_quantile(p):
    """Inverse normal CDF (Acklam rational approximation, accurate to ~4 d.p.).

    Used for the Q-Q plot so scipy is not a hard dependency.
    Reference: Peter J. Acklam, acklam.net/stats/distrib/norm/normsinv.html.
    """
    a = [-3.969683028665376e+01,  2.209460984245205e+02,
         -2.759285104469687e+02,  1.383577518672690e+02,
         -3.066479806614716e+01,  2.506628277459239e+00]
    b = [-5.447609879822406e+01,  1.615858368580409e+02,
         -1.556989798598866e+02,  6.680131188771972e+01,
         -1.328068155288572e+01]
    c = [-7.784894002430293e-03, -3.223964580411365e-01,
         -2.400758277161838e+00, -2.549732539343734e+00,
          4.374664141464968e+00,  2.938163982698783e+00]
    d = [ 7.784695709041462e-03,  3.224671290700398e-01,
          2.445134137142996e+00,  3.754408661907416e+00]
    p_lo = 0.02425
    p_hi = 1 - p_lo
    if p < p_lo:
        q = math.sqrt(-2 * math.log(p))
        return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / \
               ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1)
    if p <= p_hi:
        q = p - 0.5
        r = q * q
        return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / \
               (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1)
    q = math.sqrt(-2 * math.log(1 - p))
    return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / \
              ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1)


def _stamp(fig, page_no, n_pages, report_id, date_str, title):
    """Stamp a header rule + footer (report id / date / page x of y) on a figure.

    Must be called AFTER axes content is built.  Header and footer live in
    figure-fraction coordinates; keep all axes content within y=[0.07, 0.94]
    to avoid collision.
    """
    from matplotlib.lines import Line2D

    # Header: centered report title above a thin rule
    fig.text(0.50, 0.970, title,
             ha="center", va="top", fontsize=8, color=_C_MUTED,
             transform=fig.transFigure)
    fig.add_artist(
        Line2D([0.07, 0.93], [0.958, 0.958],
               color=_C_GRID, lw=0.8, transform=fig.transFigure, clip_on=False)
    )

    # Footer: thin rule then three label fields
    fig.add_artist(
        Line2D([0.07, 0.93], [0.044, 0.044],
               color=_C_GRID, lw=0.8, transform=fig.transFigure, clip_on=False)
    )
    fig.text(0.07, 0.033, f"Report: {report_id}",
             ha="left", va="top", fontsize=7, color=_C_MUTED, transform=fig.transFigure)
    fig.text(0.50, 0.033, date_str,
             ha="center", va="top", fontsize=7, color=_C_MUTED, transform=fig.transFigure)
    fig.text(0.93, 0.033, f"Page {page_no} of {n_pages}",
             ha="right", va="top", fontsize=7, color=_C_MUTED, transform=fig.transFigure)


def _text_table(ax, data, col_labels, col_x, row_y_start, row_height=0.062, fontsize=8.5):
    """Render a booktabs-style text table on an axis with data coords [0,1]x[0,1].

    Draws top rule, header underline (midrule), data rows, and bottom rule using
    horizontal lines only -- no vertical rules, no cell fills.  Numbers in all
    columns except the first are right-aligned (at the right edge of the column).
    The Status column (if col_labels[-1] == 'Status') has its value colored.

    Args:
        ax:           Axes with axis("off"), xlim=[0,1], ylim=[0,1].
        data:         List of row lists (strings).
        col_labels:   Column header strings.
        col_x:        Left x edge of each column (data coords).  The right edge
                      of column j is col_x[j+1] (or 1.0 for the last column).
        row_y_start:  y coordinate (data coords) of the top rule.
        row_height:   Row height in data coords.
        fontsize:     Font size for all table text.

    Returns:
        y coordinate of the bottom rule.
    """
    n_cols = len(col_labels)
    # Right edge of each column: left edge of the next column.
    # For the last column, extrapolate from the two rightmost left edges but cap
    # at 0.98 so it never renders outside the content area regardless of how many
    # columns are in the table.
    last_right = min(0.98, col_x[-1] + (col_x[-1] - col_x[-2]))
    col_right = list(col_x[1:]) + [last_right]

    def rule(y_r, lw=0.7):
        ax.plot([col_x[0], col_right[-1]], [y_r, y_r],
                color=_C_TEXT, lw=lw, transform=ax.transData, clip_on=False)

    y = row_y_start
    rule(y, lw=1.0)   # toprule
    y -= row_height * 0.2

    # Header row
    for j, label in enumerate(col_labels):
        x = col_x[j] if j == 0 else col_right[j] - 0.018
        ha = "left" if j == 0 else "right"
        ax.text(x, y - row_height * 0.45, label,
                ha=ha, va="center", fontsize=fontsize, fontweight="bold",
                color=_C_TEXT)
    y -= row_height
    rule(y, lw=0.5)   # midrule

    # Data rows
    is_status_col = col_labels[-1] == "Status"
    for row in data:
        y_text = y - row_height * 0.5
        for j, cell in enumerate(row):
            color = _C_TEXT
            if is_status_col and j == n_cols - 1:
                if "PASS" in cell or cell == "OK":
                    color = _C_PASS
                elif "WARN" in cell:
                    color = _C_WARN
                elif "FAIL" in cell:
                    color = _C_FAIL
            # Right-aligned columns use a 0.018 inset from the column right edge
            # to create a visible gap between adjacent columns.
            x = col_x[j] if j == 0 else col_right[j] - 0.018
            ha = "left" if j == 0 else "right"
            ax.text(x, y_text, cell, ha=ha, va="center",
                    fontsize=fontsize, color=color)
        y -= row_height

    rule(y, lw=1.0)   # bottomrule
    return y


# ---------------------------------------------------------------------------
# PDF report: main generator
# ---------------------------------------------------------------------------

def generate_report(csv_path, min_power, rows, coeffs, residuals, channels,
                    write_png=False):
    """Write a professional 4-page PDF calibration report next to the input CSV.

    Pages:
        1  Cover / summary -- title block, fitted model equation, verdict and
           key fit statistics, firmware commands, acceptance guidance.
        2  Model and statistics -- coefficient table with standard errors and
           95% confidence intervals, goodness-of-fit table, conditioning note.
        3  Residual diagnostics -- parity plot (hero), residuals vs fitted /
           vs frequency / vs sqrt(P), normal Q-Q plot.
        4  Calibration surface and per-frequency small multiples -- volt(f,P)
           contour, calibration-benefit deviation panel, 3x5 per-frequency grid.

    Args:
        csv_path:   Input CSV path (report saved alongside it).
        min_power:  Minimum power filter applied when loading (for labelling).
        rows:       Loaded data rows: list of (freq_mhz, keysight_w, volt).
        coeffs:     Fitted polynomial coefficients [a, b, c, d, e].
        residuals:  Fit residuals (predicted - measured).
        channels:   RFG channel numbers programmed (e.g. [1] or [1, 2]).
        write_png:  If True, also save each page as a PNG beside the PDF.

    Returns:
        Path to the generated PDF, or None if matplotlib is unavailable.
    """
    try:
        import matplotlib
        import matplotlib.pyplot as plt
        import matplotlib.colors as mcolors
        import matplotlib.gridspec as gridspec
        from matplotlib.backends.backend_pdf import PdfPages
    except ImportError:
        print(
            "matplotlib not installed -- skipping report.  "
            "Install with: pip install matplotlib",
            file=sys.stderr,
        )
        return None

    base = os.path.splitext(os.path.abspath(csv_path))[0]
    # Try the canonical name first; fall back to _v2, _v3 etc. if the file is
    # locked (e.g. open in a PDF viewer).
    report_path = base + "_report.pdf"
    for _suffix in ("", "_v2", "_v3", "_v4"):
        candidate = base + _suffix + "_report.pdf"
        try:
            with open(candidate, "ab"):
                pass
            report_path = candidate
            break
        except PermissionError:
            pass

    a, b, c, d, e = coeffs
    sentinel_ok = SENTINEL_C_MIN < c < SENTINEL_C_MAX

    A_mat, y_vec = build_matrix(rows)
    predicted = A_mat @ coeffs
    res = predicted - y_vec

    stats = compute_fit_statistics(rows, coeffs, res)

    freqs_sorted = sorted(set(r[0] for r in rows))
    n_freqs = len(freqs_sorted)
    freq_stats = _per_freq_stats(rows, res)
    cmd_lines = build_commands(coeffs, channels)
    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d %H:%M")

    report_id = os.path.basename(base)
    ch_label  = ", ".join(f"CH_{'A' if ch == 1 else 'B'}" for ch in channels)
    report_title = "Pulse-Mode Drive Calibration Report"

    # Frequency-indexed colormap for scatter plots
    freq_idx = {f: i for i, f in enumerate(freqs_sorted)}
    cmap_seq = plt.cm.viridis
    point_colors = [cmap_seq(freq_idx[r[0]] / max(n_freqs - 1, 1)) for r in rows]
    sm_freq = plt.cm.ScalarMappable(
        cmap=cmap_seq,
        norm=mcolors.Normalize(vmin=freqs_sorted[0], vmax=freqs_sorted[-1]),
    )
    sm_freq.set_array([])

    # Calibration surface grid (shared by page 4)
    f_lo, f_hi = freqs_sorted[0], freqs_sorted[-1]
    p_lo = max(0.5, min(r[1] for r in rows) * 0.80)
    p_hi = max(r[1] for r in rows) * 1.05
    F_surf, P_surf = np.meshgrid(np.linspace(f_lo, f_hi, 100),
                                  np.linspace(p_lo, p_hi, 80))
    V_poly = (a + b * F_surf + c * np.sqrt(P_surf)
              + d * F_surf**2 + e * F_surf * np.sqrt(P_surf))

    coeff_names = ["a", "b", "c", "d", "e"]
    coeff_units = ["V", "V/MHz", "V/W^1/2", "V/MHz^2", "V/(MHz W^1/2)"]

    figs = []   # collect all figures before stamping so we know n_pages

    with matplotlib.rc_context(_REPORT_RC):

        # ------------------------------------------------------------------ #
        # Page 1 -- Cover / summary                                           #
        # ------------------------------------------------------------------ #
        fig = plt.figure(figsize=(8.5, 11))
        fig.patch.set_facecolor("white")
        # Manual axes: all content between y=0.07 and y=0.94 (stamp occupies outside)
        ax = fig.add_axes([0.08, 0.07, 0.84, 0.87])
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")

        y = 0.985

        # Title
        ax.text(0.5, y, report_title,
                ha="center", va="top", fontsize=15, fontweight="bold",
                color=_C_TEXT)
        y -= 0.055
        ax.plot([0, 1], [y, y], color=_C_ACCENT, lw=1.5, clip_on=False)
        y -= 0.030

        # Metadata (two-column layout)
        meta = [
            ("Report ID",       report_id,
             "Source log",      os.path.basename(csv_path)),
            ("Channel(s)",      ch_label,
             "Date / time",     now_str),
            ("Points fitted",   f"{stats['n']}  (min power: {min_power} W)",
             "Frequency range", f"{freqs_sorted[0]:.3f}–{freqs_sorted[-1]:.3f} MHz  ({n_freqs} steps)"),
        ]
        for lk, lv, rk, rv in meta:
            ax.text(0.00, y, lk + ":", ha="left", va="top", fontsize=8.5, color=_C_MUTED)
            ax.text(0.18, y, lv,       ha="left", va="top", fontsize=8.5, color=_C_TEXT)
            ax.text(0.52, y, rk + ":", ha="left", va="top", fontsize=8.5, color=_C_MUTED)
            ax.text(0.72, y, rv,       ha="left", va="top", fontsize=8.5, color=_C_TEXT)
            y -= 0.038
        y -= 0.012
        ax.plot([0, 1], [y, y], color=_C_GRID, lw=0.6, clip_on=False)
        y -= 0.028

        # Fitted model equation
        ax.text(0.0, y, "Fitted model:", ha="left", va="top", fontsize=8.5, color=_C_MUTED)
        y -= 0.048
        eq = (r"$\mathrm{volt} = a + b\,f + c\,\sqrt{P}"
              r" + d\,f^2 + e\,f\,\sqrt{P}$"
              "\n"
              r"$f=\mathrm{freq\,(MHz)},\quad P=\mathrm{Keysight\,power\,(W)},"
              r"\quad \mathrm{volt}=\sqrt{\mathrm{drive\_sum}}\times"
              rf"{DRIVE_VOLT_SCALE:.4f}$")
        ax.text(0.03, y, eq, ha="left", va="top", fontsize=10, color=_C_TEXT)
        y -= 0.115
        ax.plot([0, 1], [y, y], color=_C_GRID, lw=0.6, clip_on=False)
        y -= 0.028

        # Verdict
        ax.text(0.0, y, "Calibration verdict:", ha="left", va="top",
                fontsize=8.5, color=_C_MUTED)
        y -= 0.040

        sentinel_color = _C_PASS if sentinel_ok else _C_FAIL
        sentinel_label = "PASS" if sentinel_ok else "FAIL"
        ax.text(0.0, y, f"Firmware sentinel  (0.01 < c < 10.0):",
                ha="left", va="top", fontsize=9, color=_C_TEXT)
        ax.text(0.58, y, f"{sentinel_label}   (c = {c:.6f})",
                ha="left", va="top", fontsize=9, fontweight="bold",
                color=sentinel_color)
        y -= 0.042

        # Headline statistics (two-column)
        m_left = [
            ("SEE  (s, residual std dev)",  f"{stats['s']:.4f} V"),
            ("Max absolute residual",       f"{stats['max_resid']:.4f} V"),
            ("Mean residual (bias)",        f"{stats['bias']:+.4f} V"),
        ]
        m_right = [
            ("RMS power error (mid-range)", f"~{stats['power_err_rms_pct']:.1f}%"),
            ("R² adjusted",            f"{stats['R2_adj']:.6f}"),
            ("Expanded U  (k=2, ~95%)",     f"±{stats['U_k2']:.4f} V"),
        ]
        for (lk, lv), (rk, rv) in zip(m_left, m_right):
            ax.text(0.02, y, lk + ":", ha="left", va="top", fontsize=8.5, color=_C_MUTED)
            ax.text(0.38, y, lv,       ha="left", va="top", fontsize=8.5, color=_C_TEXT)
            ax.text(0.55, y, rk + ":", ha="left", va="top", fontsize=8.5, color=_C_MUTED)
            ax.text(0.84, y, rv,       ha="right", va="top", fontsize=8.5, color=_C_TEXT)
            y -= 0.038
        y -= 0.012
        ax.plot([0, 1], [y, y], color=_C_GRID, lw=0.6, clip_on=False)
        y -= 0.028

        # Firmware commands
        ax.text(0.0, y, "Firmware commands  (RFG ASCII port, 256000 baud):",
                ha="left", va="top", fontsize=8.5, color=_C_MUTED)
        y -= 0.038
        for cmd in cmd_lines:
            ax.text(0.03, y, cmd, ha="left", va="top", fontsize=8.5,
                    color=_C_TEXT, fontfamily="monospace")
            y -= 0.030
        y -= 0.008
        ax.plot([0, 1], [y, y], color=_C_GRID, lw=0.6, clip_on=False)
        y -= 0.025

        # Acceptance guidance (footnote style)
        guidance = (
            f"Acceptance guidance:  RMS residual < {THRESHOLD_TARGET_V:.2f} V = target "
            f"(~3–5% power error) |  < {THRESHOLD_LIMIT_V:.2f} V = acceptable |"
            f"  > {THRESHOLD_LIMIT_V:.2f} V = re-sweep (try --min-power 3)."
        )
        note_u = (
            "Uncertainty U = 2s covers only the model Type-A lack-of-fit term.  "
            "A full budget combines Type-B contributions from the Keysight N1914A "
            "and signal generator in quadrature."
        )
        for line in (guidance, note_u):
            ax.text(0.0, y, line, ha="left", va="top", fontsize=7.5, color=_C_MUTED)
            y -= 0.036

        figs.append(fig)

        # ------------------------------------------------------------------ #
        # Page 2 -- Model coefficients and fit statistics                     #
        # ------------------------------------------------------------------ #
        fig = plt.figure(figsize=(8.5, 11))
        fig.patch.set_facecolor("white")
        ax = fig.add_axes([0.08, 0.07, 0.84, 0.87])
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")

        y = 0.985
        ax.text(0.5, y, "Model Coefficients and Fit Statistics",
                ha="center", va="top", fontsize=13, fontweight="bold", color=_C_TEXT)
        y -= 0.048
        ax.plot([0, 1], [y, y], color=_C_ACCENT, lw=1.2, clip_on=False)
        y -= 0.030

        # -- Table 1: coefficient estimates with uncertainties --
        ax.text(0.0, y, "Table 1.  Fitted polynomial coefficients  "
                r"(volt = a + b·f + c·√P + d·f² + e·f·√P)",
                ha="left", va="top", fontsize=9, fontweight="bold", color=_C_TEXT)
        y -= 0.030

        # Column positions (left edge x) and headers
        c_x  = [0.00, 0.07, 0.22, 0.38, 0.53, 0.68, 0.83]
        c_hdr = ["Term", "Units", "Estimate", "Std error", "CI lo (95%)", "CI hi (95%)", "t-value"]
        c_dat = []
        for j in range(5):
            row_d = [
                coeff_names[j],
                coeff_units[j],
                f"{coeffs[j]: .8f}",
                f"{stats['coeff_SE'][j]:.6f}",
                f"{stats['coeff_CI_lo'][j]: .6f}",
                f"{stats['coeff_CI_hi'][j]: .6f}",
                f"{stats['coeff_t'][j]:.2f}",
            ]
            c_dat.append(row_d)

        y = _text_table(ax, c_dat, c_hdr, c_x,
                        row_y_start=y, row_height=0.060, fontsize=8.0)
        y -= 0.012
        ax.text(0.0, y,
                f"95% CI uses t_crit = {stats['t_crit']:.4f}  "
                f"(dof = {stats['dof']}, two-tailed).",
                ha="left", va="top", fontsize=7.5, color=_C_MUTED)
        y -= 0.040

        ax.plot([0, 1], [y, y], color=_C_GRID, lw=0.5, clip_on=False)
        y -= 0.030

        # -- Table 2: goodness of fit --
        ax.text(0.0, y, "Table 2.  Goodness-of-fit statistics",
                ha="left", va="top", fontsize=9, fontweight="bold", color=_C_TEXT)
        y -= 0.030

        # 2-column layout: Statistic | Value
        # Notes / explanations are given in the paragraph that follows the table.
        g_x   = [0.00, 0.68]
        g_hdr = ["Statistic", "Value"]
        g_dat = [
            ["SEE  (residual std dev, s)",          f"{stats['s']:.6f} V"],
            ["RMSE  (denominator n)",               f"{stats['RMSE']:.6f} V"],
            ["R²",                             f"{stats['R2']:.6f}"],
            ["R² adjusted",                    f"{stats['R2_adj']:.6f}"],
            ["Max |residual|",                      f"{stats['max_resid']:.6f} V"],
            ["Mean residual (bias)",                f"{stats['bias']:+.6f} V"],
            ["n  /  p  /  dof",
             f"{stats['n']}  /  {stats['p']}  /  {stats['dof']}"],
            ["F-statistic  (overall model)",        f"{stats['F_stat']:.1f}"],
            ["Expanded uncertainty U  (k=2, ~95%)", f"±{stats['U_k2']:.4f} V"],
            ["RMS power error  (mid-range)",        f"~{stats['power_err_rms_pct']:.1f}%"],
        ]
        y = _text_table(ax, g_dat, g_hdr, g_x,
                        row_y_start=y, row_height=0.054, fontsize=8.5)
        y -= 0.018

        # Explanatory notes for Table 2
        gof_notes = (
            f"SEE = sqrt(SSE/dof): residual standard deviation -- the headline fit quality"
            f" in voltage units.  RMSE = sqrt(SSE/n).  R² adjusted penalises for term"
            f" count; prefer over R² for model comparison.  Max |residual| acceptance"
            f" limit is {THRESHOLD_LIMIT_V:.2f} V.  Bias should be ≈0 for OLS with"
            f" an intercept term.  U = 2s covers only the Type-A (lack-of-fit)"
            f" component.  Power error ≈ 2×voltage error / mean drive voltage."
        )
        ax.text(0.02, y, gof_notes, ha="left", va="top", fontsize=7.5,
                color=_C_MUTED)
        y -= 0.060

        ax.plot([0, 1], [y, y], color=_C_GRID, lw=0.5, clip_on=False)
        y -= 0.030

        # -- Design matrix conditioning --
        ax.text(0.0, y, "Design matrix conditioning",
                ha="left", va="top", fontsize=9, fontweight="bold", color=_C_TEXT)
        y -= 0.035
        cond_ok = stats['cond_number'] < 30
        vif_ok  = stats['max_vif'] < 10
        cond_lines = [
            (f"Condition number (column-normalised design matrix):  "
             f"{stats['cond_number']:.1f}  "
             f"-- {'acceptable' if cond_ok else 'ELEVATED -- collinear terms; individual CIs are wide'}"),
            (f"Max VIF (non-intercept predictors):  "
             f"{stats['max_vif']:.1f}  "
             f"-- {'acceptable (< 10)' if vif_ok else 'ELEVATED > 10 -- multicollinearity present'}"),
            ("The model includes correlated terms (f and f²; √P and f·√P).  "
             "A condition number up to ~100 is expected for this polynomial family "
             "and does not invalidate predictions -- it widens coefficient CIs."),
        ]
        for line in cond_lines:
            ax.text(0.02, y, line, ha="left", va="top", fontsize=8.5,
                    color=_C_TEXT if "ELEVATED" not in line else _C_WARN)
            y -= 0.040

        figs.append(fig)

        # ------------------------------------------------------------------ #
        # Page 3 -- Residual diagnostics                                      #
        # ------------------------------------------------------------------ #
        fig = plt.figure(figsize=(8.5, 11))
        fig.patch.set_facecolor("white")
        gs3 = gridspec.GridSpec(
            3, 2, figure=fig,
            left=0.10, right=0.94, top=0.92, bottom=0.07,
            hspace=0.55, wspace=0.40,
        )

        volt_meas = y_vec.tolist()
        volt_pred = predicted.tolist()
        pwr_all   = [r[1] for r in rows]
        frq_all   = [r[0] for r in rows]
        sqP_all   = [math.sqrt(r[1]) for r in rows]
        s_fit     = stats['s']
        tol_2s    = 2.0 * s_fit

        # Parity plot -- hero figure spanning full top row
        ax_p = fig.add_subplot(gs3[0, :])
        ax_p.scatter(volt_meas, volt_pred, c=point_colors,
                     s=20, alpha=0.75, linewidths=0, zorder=3)
        v_lo = min(volt_meas + volt_pred)
        v_hi = max(volt_meas + volt_pred)
        ax_p.plot([v_lo, v_hi], [v_lo, v_hi],
                  color=_C_MUTED, lw=0.9, ls="--", label="y = x  (ideal)", zorder=2)
        ax_p.fill_between([v_lo, v_hi],
                          [v_lo - tol_2s, v_hi - tol_2s],
                          [v_lo + tol_2s, v_hi + tol_2s],
                          color=_C_ACCENT, alpha=0.07,
                          label=f"±2s = ±{tol_2s:.4f} V")
        ax_p.set_aspect("equal")
        ax_p.set_xlabel("Measured voltage (V)")
        ax_p.set_ylabel("Predicted voltage (V)")
        ax_p.set_title("Parity: Predicted vs Measured")
        ax_p.legend(loc="upper left")
        cb = fig.colorbar(sm_freq, ax=ax_p, fraction=0.022, pad=0.02)
        cb.set_label("Frequency (MHz)", fontsize=8)
        ax_p.text(0.97, 0.05,
                  f"$R^2_{{\\mathrm{{adj}}}}$ = {stats['R2_adj']:.5f}\n"
                  f"RMSE = {stats['RMSE']:.4f} V",
                  transform=ax_p.transAxes, ha="right", va="bottom", fontsize=8,
                  bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=_C_GRID, lw=0.7))

        def _resid_panel(ax_r, x_vals, xlabel, title):
            ax_r.scatter(x_vals, res, c=point_colors, s=12, alpha=0.7, linewidths=0)
            ax_r.axhline(0, color=_C_MUTED, lw=0.8, ls="--")
            ax_r.axhspan(-THRESHOLD_TARGET_V, THRESHOLD_TARGET_V,
                         color=_C_PASS, alpha=0.07,
                         label=f"±{THRESHOLD_TARGET_V:.2f} V")
            ax_r.set_xlabel(xlabel)
            ax_r.set_ylabel("Residual (V)")
            ax_r.set_title(title)
            ax_r.legend(fontsize=7, loc="upper right")

        _resid_panel(fig.add_subplot(gs3[1, 0]),
                     volt_pred, "Predicted voltage (V)", "Residuals vs Fitted")
        _resid_panel(fig.add_subplot(gs3[1, 1]),
                     frq_all, "Frequency (MHz)", "Residuals vs Frequency")
        _resid_panel(fig.add_subplot(gs3[2, 0]),
                     sqP_all, r"$\sqrt{P}$  (W$^{1/2}$)", r"Residuals vs $\sqrt{P}$")

        # Normal Q-Q plot of standardised residuals
        ax_qq = fig.add_subplot(gs3[2, 1])
        res_std = res / s_fit if s_fit > 0 else res
        res_std_sorted = np.sort(res_std)
        n_r = len(res_std_sorted)
        th_q = np.array([_normal_quantile((i + 0.5) / n_r) for i in range(n_r)])
        ax_qq.scatter(th_q, res_std_sorted, color=_C_ACCENT, s=12, alpha=0.70, linewidths=0)
        lim = max(abs(th_q[0]), abs(th_q[-1])) * 1.05
        ax_qq.plot([-lim, lim], [-lim, lim],
                   color=_C_MUTED, lw=0.9, ls="--", label="normal reference")
        ax_qq.set_xlabel("Theoretical standard-normal quantile")
        ax_qq.set_ylabel("Standardised residual")
        ax_qq.set_title("Normal Q-Q Plot of Residuals")
        ax_qq.legend(fontsize=7)

        figs.append(fig)

        # ------------------------------------------------------------------ #
        # Page 4 -- Calibration surface and per-frequency small multiples    #
        # ------------------------------------------------------------------ #
        fig = plt.figure(figsize=(8.5, 11))
        fig.patch.set_facecolor("white")
        gs4 = gridspec.GridSpec(
            2, 1, figure=fig,
            left=0.09, right=0.94, top=0.92, bottom=0.07,
            hspace=0.48, height_ratios=[1.0, 1.35],
        )

        # -- Top: two surface panels --
        gs4_top = gridspec.GridSpecFromSubplotSpec(
            1, 2, subplot_spec=gs4[0], wspace=0.42)

        ax_surf = fig.add_subplot(gs4_top[0])
        cf = ax_surf.contourf(F_surf, P_surf, V_poly, levels=12, cmap="viridis")
        ax_surf.contour(F_surf, P_surf, V_poly,
                        levels=6, colors="white", linewidths=0.3, alpha=0.35)
        ax_surf.scatter([r[0] for r in rows], [r[1] for r in rows],
                        c="white", s=8, alpha=0.45, linewidths=0, label="measured")
        cb = fig.colorbar(cf, ax=ax_surf, fraction=0.046, pad=0.03)
        cb.set_label("Drive voltage (V)", fontsize=8)
        ax_surf.set_xlabel("Frequency (MHz)")
        ax_surf.set_ylabel("Power setpoint (W)")
        ax_surf.set_title("Drive Voltage  volt(f, P)")
        ax_surf.legend(fontsize=7, loc="upper left")

        V_uncal = DRIVE_VOLT_SCALE * np.sqrt(P_surf)
        pdev_pct = ((V_poly / V_uncal) ** 2 - 1.0) * 100.0
        abs_lim = max(5.0, float(np.percentile(np.abs(pdev_pct), 98)))
        norm_div = mcolors.TwoSlopeNorm(vcenter=0, vmin=-abs_lim, vmax=abs_lim)

        ax_ben = fig.add_subplot(gs4_top[1])
        cf2 = ax_ben.contourf(F_surf, P_surf, pdev_pct,
                               levels=12, cmap="RdBu_r", norm=norm_div)
        ax_ben.contour(F_surf, P_surf, pdev_pct,
                       levels=[0], colors=_C_TEXT, linewidths=0.6, alpha=0.5)
        ax_ben.scatter([r[0] for r in rows], [r[1] for r in rows],
                       c=_C_TEXT, s=8, alpha=0.25, linewidths=0)
        cb2 = fig.colorbar(cf2, ax=ax_ben, fraction=0.046, pad=0.03)
        cb2.set_label("Power correction vs uncal (%)", fontsize=7)
        ax_ben.set_xlabel("Frequency (MHz)")
        ax_ben.set_ylabel("Power setpoint (W)")
        ax_ben.set_title("Calibration Benefit\n"
                          "(red = drives harder, blue = softer than uncalibrated)")

        # -- Bottom: per-frequency small multiples (Trellis, 5 columns) --
        n_cols_sm = 5
        n_rows_sm = math.ceil(n_freqs / n_cols_sm)
        gs4_bot = gridspec.GridSpecFromSubplotSpec(
            n_rows_sm, n_cols_sm, subplot_spec=gs4[1],
            hspace=0.60, wspace=0.42)

        y_sm_lim = max(THRESHOLD_LIMIT_V * 1.35, float(np.max(np.abs(res))) * 1.20)
        p_lo_sm = min(r[1] for r in rows) * 0.85
        p_hi_sm = max(r[1] for r in rows) * 1.10

        # Map freq -> [(power, residual)] for panel fill
        freq_data = {}
        for i, row in enumerate(rows):
            freq_data.setdefault(row[0], []).append((row[1], float(res[i])))

        for idx, freq in enumerate(freqs_sorted):
            ri = idx // n_cols_sm
            ci = idx % n_cols_sm
            ax_sm = fig.add_subplot(gs4_bot[ri, ci])

            pwr_v = [p for p, _ in freq_data[freq]]
            res_v = [rv for _, rv in freq_data[freq]]

            ax_sm.scatter(pwr_v, res_v, color=_C_ACCENT,
                          s=10, alpha=0.80, linewidths=0, zorder=3)
            ax_sm.axhline(0, color=_C_MUTED, lw=0.5, ls="-")
            ax_sm.axhspan(-THRESHOLD_TARGET_V, THRESHOLD_TARGET_V,
                          color=_C_PASS, alpha=0.09)
            ax_sm.set_xlim(p_lo_sm, p_hi_sm)
            ax_sm.set_ylim(-y_sm_lim, y_sm_lim)
            ax_sm.set_title(f"{freq:.3f} MHz", fontsize=7, pad=2)
            ax_sm.spines["top"].set_visible(False)
            ax_sm.spines["right"].set_visible(False)
            ax_sm.tick_params(axis="both", labelsize=5.5)
            ax_sm.yaxis.grid(True, color=_C_GRID, lw=0.4, alpha=0.7)

            # Status symbol (top-right corner of panel)
            f_rms = next(s_f[2] for s_f in freq_stats if s_f[0] == freq)
            if f_rms < THRESHOLD_TARGET_V:
                sym, sc = "●", _C_PASS
            elif f_rms < THRESHOLD_LIMIT_V:
                sym, sc = "▲", _C_WARN
            else:
                sym, sc = "■", _C_FAIL
            ax_sm.text(0.96, 0.96, sym, transform=ax_sm.transAxes,
                       ha="right", va="top", fontsize=7, color=sc)

            # Axis labels only on edge panels to reduce clutter
            if ri == n_rows_sm - 1:
                ax_sm.set_xlabel("P (W)", fontsize=6)
            else:
                ax_sm.set_xticklabels([])
            if ci == 0:
                ax_sm.set_ylabel("Resid (V)", fontsize=6)
            else:
                ax_sm.set_yticklabels([])

        # Hide unused grid cells when n_freqs < n_rows_sm * n_cols_sm
        for spare in range(n_freqs, n_rows_sm * n_cols_sm):
            fig.add_subplot(gs4_bot[spare // n_cols_sm, spare % n_cols_sm]).set_visible(False)

        # Section label for the small-multiples block
        fig.text(0.515, 0.476,
                 f"Per-frequency residuals  "
                 f"(● < {THRESHOLD_TARGET_V:.2f} V target   "
                 f"▲ < {THRESHOLD_LIMIT_V:.2f} V acceptable   "
                 f"■ = re-sweep)",
                 ha="center", va="center", fontsize=7.5, color=_C_MUTED)

        figs.append(fig)

    # ---------------------------------------------------------------------- #
    # Stamp all pages and write to PDF                                        #
    # ---------------------------------------------------------------------- #
    n_pages = len(figs)
    with PdfPages(report_path) as pdf:
        d = pdf.infodict()
        d["Title"]        = report_title
        d["Author"]       = "fit_drive_cal.py"
        d["Subject"]      = f"RFG drive calibration -- {report_id}"
        d["Keywords"]     = "RFG, calibration, drive polynomial, TVN"
        d["CreationDate"] = now
        d["ModDate"]      = now

        for page_no, fig in enumerate(figs, start=1):
            _stamp(fig, page_no, n_pages, report_id, now_str, report_title)
            pdf.savefig(fig)   # no bbox_inches -- uniform 8.5 x 11 geometry
            if write_png:
                png_path = f"{base}_report_p{page_no}.png"
                fig.savefig(png_path, dpi=150)
                print(f"PNG page {page_no}: {png_path}")
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

    rows = load_csv(args.csv_file, args.min_power, channel_filter=args.channel)
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
        generate_report(args.csv_file, args.min_power, rows, coeffs, residuals,
                        channels, write_png=args.png)

    if args.port:
        send_commands(args.port, args.baud, commands)

    if args.plot:
        plot_results(coeffs, residuals, rows)


if __name__ == "__main__":
    main()
