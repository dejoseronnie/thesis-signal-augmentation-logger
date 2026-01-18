import argparse
import csv
import datetime as dt
import os
import re
import subprocess
import threading
import time
from dataclasses import dataclass
from typing import Optional, List, Tuple, Dict

import openpyxl

try:
    import serial  # pyserial
except ImportError:
    serial = None


# ----------------------------- GPS -----------------------------

def nmea_ddmm_to_decimal(ddmm: str, direction: str) -> Optional[float]:
    if not ddmm or not direction:
        return None
    try:
        if len(ddmm.split(".")[0]) <= 4:
            deg = int(ddmm[0:2])
            minutes = float(ddmm[2:])
        else:
            deg = int(ddmm[0:3])
            minutes = float(ddmm[3:])
        dec = deg + minutes / 60.0
        if direction in ("S", "W"):
            dec = -dec
        return dec
    except Exception:
        return None


def parse_nmea_gga(line: str) -> Optional[Tuple[Optional[float], Optional[float]]]:
    # $GPGGA,hhmmss.ss,llll.ll,a,yyyyy.yy,a,x,xx,x.x,x.x,M,...
    if not (line.startswith("$GPGGA") or line.startswith("$GNGGA")):
        return None
    parts = line.strip().split(",")
    if len(parts) < 6:
        return None
    fixq = parts[6].strip() if len(parts) > 6 else "0"
    if fixq == "0" or fixq == "":
        return (None, None)

    lat_raw, lat_dir = parts[2].strip(), parts[3].strip()
    lon_raw, lon_dir = parts[4].strip(), parts[5].strip()
    lat = nmea_ddmm_to_decimal(lat_raw, lat_dir)
    lon = nmea_ddmm_to_decimal(lon_raw, lon_dir)
    return (lat, lon)


def get_gps_fix_from_serial(port: str, baud: int, timeout_s: float = 0.5) -> Tuple[Optional[float], Optional[float]]:
    """
    Reads NMEA from a COM port (e.g., COM3). Returns (lat, lon) or (None, None).
    """
    if serial is None:
        return (None, None)

    try:
        with serial.Serial(port, baudrate=baud, timeout=timeout_s) as ser:
            t0 = time.time()
            while time.time() - t0 < 1.0:
                line = ser.readline().decode(errors="ignore").strip()
                parsed = parse_nmea_gga(line)
                if parsed is None:
                    continue
                lat, lon = parsed
                if lat is not None and lon is not None:
                    return (lat, lon)
    except Exception:
        return (None, None)

    return (None, None)


# ----------------------------- WIFI (Windows via netsh) -----------------------------

def run_netsh_show_interfaces() -> str:
    # netsh output is localized on some systems; this works best on English Windows.
    # If your Windows is not English, tell me and I'll adapt the parser.
    out = subprocess.check_output(["netsh", "wlan", "show", "interfaces"], stderr=subprocess.STDOUT)
    return out.decode(errors="replace")


def parse_signal_percent(netsh_text: str) -> Optional[int]:
    # Example line: "    Signal                 : 86%"
    m = re.search(r"Signal\s*:\s*(\d+)\s*%", netsh_text)
    if not m:
        return None
    return int(m.group(1))


def approx_rssi_dbm_from_percent(signal_percent: int) -> int:
    # Common approximation used in practice:
    # RSSI(dBm) ≈ (Signal% / 2) - 100
    return int(round(signal_percent / 2.0 - 100.0))


# ----------------------------- CLASSIFICATION -----------------------------

def classify_rssi(rssi_dbm: Optional[int]) -> str:
    # You can change these to match your thesis bins.
    if rssi_dbm is None:
        return "Unknown"
    if rssi_dbm >= -60:
        return "Excellent"
    if rssi_dbm >= -70:
        return "Good"
    if rssi_dbm >= -80:
        return "Fair"
    if rssi_dbm >= -90:
        return "Weak"
    return "Very Weak"


# ----------------------------- DATA MODELS -----------------------------

@dataclass
class RawRow:
    sample: int
    gps: str               # "lat,lon" or ""
    rssi_dbm: Optional[int]
    snr_db: Optional[int]
    timestamp: str         # ISO


# ----------------------------- GROUPING / SUMMARY -----------------------------

@dataclass
class SummaryRow:
    group_label: str
    total_points: int
    coverage_percent: float
    snr_mean: Optional[float]
    pass_fail: str


def cluster_by_rssi_tolerance(values: List[int], tol_db: int) -> List[List[int]]:
    """
    Clusters RSSI values into groups where adjacent sorted values differ by <= tol_db.
    Example: tol_db=3 groups [-83,-82,-81] together.
    """
    if not values:
        return []
    vals = sorted(values)
    clusters = [[vals[0]]]
    for v in vals[1:]:
        if abs(v - clusters[-1][-1]) <= tol_db:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return clusters


def compute_summary(raw_rows: List[RawRow], tol_db: int, usable_threshold_dbm: int, pass_threshold_pct: float) -> List[SummaryRow]:
    """
    Groups rows by RSSI similarity (within tolerance), then computes:
    - total points in group
    - coverage% (fraction of points in group with RSSI >= usable_threshold)
    - mean SNR in group
    - pass/fail by coverage% >= pass_threshold_pct
    """
    rssi_list = [r.rssi_dbm for r in raw_rows if r.rssi_dbm is not None]
    rssi_vals = [v for v in rssi_list if v is not None]
    clusters = cluster_by_rssi_tolerance(rssi_vals, tol_db)

    # Map RSSI->rows (may be many rows same RSSI)
    buckets: Dict[int, List[RawRow]] = {}
    for r in raw_rows:
        if r.rssi_dbm is None:
            continue
        buckets.setdefault(r.rssi_dbm, []).append(r)

    summary: List[SummaryRow] = []
    for idx, cl in enumerate(clusters, start=1):
        # collect rows for all RSSI values in this cluster
        rows_in_cluster: List[RawRow] = []
        for v in cl:
            rows_in_cluster.extend(buckets.get(v, []))

        if not rows_in_cluster:
            continue

        total = len(rows_in_cluster)
        usable = sum(1 for r in rows_in_cluster if (r.rssi_dbm is not None and r.rssi_dbm >= usable_threshold_dbm))
        cov_pct = (usable / total) * 100.0

        snrs = [r.snr_db for r in rows_in_cluster if r.snr_db is not None]
        snr_mean = (sum(snrs) / len(snrs)) if snrs else None

        # Label the cluster by its RSSI range
        r_min = min(cl)
        r_max = max(cl)
        label = f"Group {idx}: RSSI {r_min} to {r_max} dBm (±{tol_db} dB)"

        pf = "PASS" if cov_pct >= pass_threshold_pct else "FAIL"

        summary.append(SummaryRow(
            group_label=label,
            total_points=total,
            coverage_percent=round(cov_pct, 1),
            snr_mean=(round(snr_mean, 1) if snr_mean is not None else None),
            pass_fail=pf
        ))

    return summary


# ----------------------------- OUTPUT -----------------------------

def write_raw_csv(path_csv: str, rows: List[RawRow]) -> None:
    os.makedirs(os.path.dirname(path_csv) or ".", exist_ok=True)
    with open(path_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        # exact format requested:
        w.writerow(["sample", "GPS", "WiFi Signal Strength", "SNR"])
        for r in rows:
            w.writerow([
                r.sample,
                r.gps,
                "" if r.rssi_dbm is None else r.rssi_dbm,
                "" if r.snr_db is None else r.snr_db
            ])


def write_excel_with_two_sheets(path_xlsx: str, raw_rows: List[RawRow], summary_rows: List[SummaryRow]) -> None:
    os.makedirs(os.path.dirname(path_xlsx) or ".", exist_ok=True)
    wb = openpyxl.Workbook()

    # Sheet 1: Raw
    ws1 = wb.active
    ws1.title = "Raw"
    ws1.append(["sample", "GPS", "WiFi Signal Strength (dBm)", "SNR (dB)", "timestamp"])
    for r in raw_rows:
        ws1.append([r.sample, r.gps, r.rssi_dbm, r.snr_db, r.timestamp])

    # Sheet 2: Summary
    ws2 = wb.create_sheet("Summary")
    ws2.append(["RSSI Cluster", "Total Measure Point", "Coverage Percentage (%)", "SNR (mean dB)", "Pass/Fail"])
    for s in summary_rows:
        ws2.append([s.group_label, s.total_points, s.coverage_percent, s.snr_mean, s.pass_fail])

    wb.save(path_xlsx)


# ----------------------------- INTERACTIVE LOGGER -----------------------------

class LoggerController:
    def __init__(self, gps_mode: str, gps_port: str, gps_baud: int,
                 noise_floor_dbm: int, interval_s: float):
        self.gps_mode = gps_mode
        self.gps_port = gps_port
        self.gps_baud = gps_baud
        self.noise_floor_dbm = noise_floor_dbm
        self.interval_s = interval_s

        self._collecting = threading.Event()
        self._stop_all = threading.Event()
        self._thread: Optional[threading.Thread] = None

        self.rows: List[RawRow] = []
        self.sample_counter = 0

    def start(self):
        if self._thread and self._thread.is_alive():
            print("Already running.")
            return
        self._collecting.set()
        self._stop_all.clear()
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()
        print("Started collecting...")

    def stop(self):
        self._collecting.clear()
        print("Stopped collecting.")

    def shutdown(self):
        self._collecting.clear()
        self._stop_all.set()
        if self._thread:
            self._thread.join(timeout=2)

    def _run_loop(self):
        while not self._stop_all.is_set():
            if not self._collecting.is_set():
                time.sleep(0.1)
                continue

            self.sample_counter += 1
            ts = dt.datetime.now().isoformat(timespec="seconds")

            # GPS
            lat, lon = (None, None)
            if self.gps_mode == "nmea":
                lat, lon = get_gps_fix_from_serial(self.gps_port, self.gps_baud)
            gps_str = f"{lat:.7f},{lon:.7f}" if (lat is not None and lon is not None) else ""

            # Wi-Fi signal (% -> dBm)
            try:
                txt = run_netsh_show_interfaces()
                sig_pct = parse_signal_percent(txt)
                rssi_dbm = approx_rssi_dbm_from_percent(sig_pct) if sig_pct is not None else None
            except Exception:
                rssi_dbm = None

            # SNR
            snr_db = (rssi_dbm - self.noise_floor_dbm) if (rssi_dbm is not None) else None

            self.rows.append(RawRow(
                sample=self.sample_counter,
                gps=gps_str,
                rssi_dbm=rssi_dbm,
                snr_db=snr_db,
                timestamp=ts
            ))

            time.sleep(self.interval_s)


def main():
    ap = argparse.ArgumentParser(description="Windows Wi-Fi RSSI+GPS logger with START/STOP and 2-sheet Excel summary.")
    ap.add_argument("--outdir", default="out", help="Output directory (default: out)")
    ap.add_argument("--noise-floor", type=int, default=-95, help="Noise floor dBm for SNR (default: -95)")
    ap.add_argument("--interval", type=float, default=2.0, help="Sampling interval seconds (default: 2.0)")

    # GPS
    ap.add_argument("--gps", choices=["none", "nmea"], default="none",
                    help="GPS mode: none or nmea (serial COM) (default: none)")
    ap.add_argument("--gps-port", default="COM3", help="GPS serial port if --gps nmea (e.g., COM3)")
    ap.add_argument("--gps-baud", type=int, default=9600, help="GPS baudrate if --gps nmea (default: 9600)")

    # Summary settings
    ap.add_argument("--tol-db", type=int, default=3, choices=[1, 2, 3],
                    help="RSSI clustering tolerance in dB (1 to 3) (default: 3)")
    ap.add_argument("--usable-threshold", type=int, default=-90,
                    help="Usable RSSI threshold for coverage percentage (default: -90)")
    ap.add_argument("--pass-threshold", type=float, default=60.0,
                    help="PASS threshold for coverage percentage (default: 60.0)")

    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    raw_csv = os.path.join(args.outdir, "raw_wifi_gps_snr.csv")
    out_xlsx = os.path.join(args.outdir, "wifi_gps_results.xlsx")

    controller = LoggerController(
        gps_mode=args.gps,
        gps_port=args.gps_port,
        gps_baud=args.gps_baud,
        noise_floor_dbm=args.noise_floor,
        interval_s=args.interval
    )

    print("\nCommands: start | stop | export | quit")
    print("Tip: Let it run for your desired duration, then 'stop' and 'export'.\n")

    try:
        while True:
            cmd = input("> ").strip().lower()

            if cmd == "start":
                controller.start()

            elif cmd == "stop":
                controller.stop()

            elif cmd == "export":
                # Write raw CSV
                write_raw_csv(raw_csv, controller.rows)

                # Build summary + write Excel with 2 sheets
                summary = compute_summary(
                    raw_rows=controller.rows,
                    tol_db=args.tol_db,
                    usable_threshold_dbm=args.usable_threshold,
                    pass_threshold_pct=args.pass_threshold
                )
                write_excel_with_two_sheets(out_xlsx, controller.rows, summary)

                print(f"Exported:\n- {raw_csv}\n- {out_xlsx}")

            elif cmd in ("quit", "exit"):
                break

            elif cmd == "":
                continue

            else:
                print("Unknown command. Use: start | stop | export | quit")

    finally:
        controller.shutdown()


if __name__ == "__main__":
    main()
