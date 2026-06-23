"""
volume_data.py
Generalizable pipeline for any intersection's bin-statistics data.

Usage
-----
1. Edit the CONFIG section below.
2. Run:  python3 volume_data.py

For each approach + movement found in the data the script produces:
  <OUTPUT_DIR>/<APPROACH>_<MOVEMENT>.csv      — hourly pivot table
  <OUTPUT_DIR>/<APPROACH>_<MOVEMENT>.xlsx     — Excel workbook with chart
"""

import csv
import os
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.legend import Legend
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG — edit this section for each new intersection
# ─────────────────────────────────────────────────────────────────────────────

INTERSECTION = "CSAH 61 (Flying Cloud Dr) at College View Dr"

# Map each approach code to the directory that holds its CSV files.
# Add or remove entries for 2- or 4-legged intersections.
APPROACH_DIRS = {
    "NB": "NB_DATA",
    "SB": "SB_DATA",
    # "EB": "EB_DATA",
    # "WB": "WB_DATA",
}

# Where to write output files (created automatically if it doesn't exist).
OUTPUT_DIR = "output"

# Movement types to process.  T=Through  L=Left  R=Right
# Remove any movements you don't need (e.g. remove "R" to skip right turns).
MOVEMENTS = ["T", "L", "R"]

MOVEMENT_LABELS = {
    "T": "Through",
    "L": "Left Turn",
    "R": "Right Turn",
}

APPROACH_LABELS = {
    "NB": "Northbound",
    "SB": "Southbound",
    "EB": "Eastbound",
    "WB": "Westbound",
}

# ─────────────────────────────────────────────────────────────────────────────
# Styling constants (match the existing _GRAPH.png palette)
# ─────────────────────────────────────────────────────────────────────────────

HOURS = list(range(23))
HOUR_LABELS = [f"{h:02d}:00" for h in HOURS]

DAY_COLORS = [
    "e6194b", "3cb44b", "4363d8", "f58231", "911eb4",
    "42d4f4", "f032e6", "a3c832", "f4a8c0", "469990",
]


# Columns used when a zone covers multiple movements (combo zones like WBTR1)
COMBO_COLUMNS = {
    "T": "ThroughCount",
    "L": "LeftTurnCount",
    "R": "RightTurnCount",
}

# ─────────────────────────────────────────────────────────────────────────────
# Data processing
# ─────────────────────────────────────────────────────────────────────────────

def discover_zones(directory, approach):
    """
    Scan all CSV files in *directory* and return a zone map:
      { zone_name: [(movement, column), ...] }

    Standard single-movement zone  → uses Volume column
      e.g. "NBT1"  → [("T", "Volume")]
           "NBL1"  → [("L", "Volume")]

    Combo multi-movement zone → uses specific count columns
      e.g. "WBTR1" → [("T", "ThroughCount"), ("R", "RightTurnCount")]
           "NBTL2" → [("T", "ThroughCount"), ("L", "LeftTurnCount")]
    """
    zone_pattern = re.compile(
        rf"^{re.escape(approach)}([TLRtlr]+)(\d+)$", re.IGNORECASE
    )
    zone_map = {}

    for fname in sorted(Path(directory).glob("*.csv")):
        with open(fname, newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                zone = row.get("ZoneName", "").strip()
                if zone in zone_map:
                    continue
                m = zone_pattern.match(zone)
                if not m:
                    continue
                letters   = m.group(1).upper()
                movements = [l for l in letters if l in MOVEMENTS]
                if not movements:
                    continue
                if len(movements) == 1:
                    # Standard zone — Volume covers the whole movement
                    zone_map[zone] = [(movements[0], "Volume")]
                else:
                    # Combo zone — each movement has its own count column
                    zone_map[zone] = [(mv, COMBO_COLUMNS[mv]) for mv in movements]

    return zone_map


def day_label(filename):
    """
    Convert a filename to a readable day label.
    E.g. 'NB_June_1.csv' -> 'June-1'
         'NB_MAY_4.csv'  -> 'May-4'
    """
    stem  = Path(filename).stem
    parts = stem.split("_")
    return f"{parts[-2].capitalize()}-{parts[-1]}"


def process_file(filepath, zone_map):
    """
    Read one CSV file and return hourly totals per movement.
    zone_map: { zone_name: [(movement, column), ...] }

    Returns: {movement: {hour: total_volume}}
    """
    hourly = defaultdict(lambda: defaultdict(int))

    with open(filepath, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            zone = row.get("ZoneName", "").strip()
            if zone not in zone_map:
                continue
            ts   = row["TimeStamp"].strip()
            hour = int(ts[11:13])
            for movement, column in zone_map[zone]:
                value = int(float(row.get(column, 0) or 0))
                hourly[movement][hour] += value

    return {mv: dict(hours) for mv, hours in hourly.items()}


def load_all_files(directory, zone_map):
    """
    Process every CSV in *directory* and return pivoted data per movement.

    Returns: {movement: {day_label: {hour: volume}}}
    """
    pivot = defaultdict(dict)

    for fname in sorted(Path(directory).glob("*.csv"),
                        key=lambda p: [int(x) if x.isdigit() else x
                                       for x in re.split(r"(\d+)", p.stem)]):
        label     = day_label(fname.name)
        file_data = process_file(fname, zone_map)
        for movement, hourly in file_data.items():
            pivot[movement][label] = hourly

    return dict(pivot)


# ─────────────────────────────────────────────────────────────────────────────
# CSV output
# ─────────────────────────────────────────────────────────────────────────────

def write_csv(pivot_data, days, col_name, out_path):
    """Write hourly pivot table + TOTAL row to a CSV file."""
    with open(out_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Hour"] + days)
        for hour, label in zip(HOURS, HOUR_LABELS):
            row = [label] + [pivot_data.get(d, {}).get(hour, 0) for d in days]
            writer.writerow(row)
        totals = [sum(pivot_data.get(d, {}).get(h, 0) for h in HOURS) for d in days]
        writer.writerow(["Total"] + totals)
    print(f"  CSV  → {out_path}")


# ─────────────────────────────────────────────────────────────────────────────
# Excel output
# ─────────────────────────────────────────────────────────────────────────────

def _thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def build_excel(pivot_data, days, title, out_path):
    """Build a styled Excel workbook with a Data sheet and a Chart sheet."""
    wb = Workbook()

    # ── Data sheet ───────────────────────────────────────────────────────────
    ws        = wb.active
    ws.title  = "Data"
    n_days    = len(days)

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_days + 1)
    tc           = ws.cell(row=1, column=1, value=title)
    tc.font      = Font(bold=True, size=13, color="FFFFFF")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    tc.fill      = PatternFill("solid", fgColor="1F4E79")
    ws.row_dimensions[1].height = 24

    # Header row
    hdr_fill = PatternFill("solid", fgColor="2E75B6")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)

    def header_cell(row, col, value):
        c           = ws.cell(row=row, column=col, value=value)
        c.font      = hdr_font
        c.fill      = hdr_fill
        c.alignment = Alignment(horizontal="center")
        c.border    = _thin_border()

    header_cell(2, 1, "Hour")
    for c, day in enumerate(days, start=2):
        header_cell(2, c, day)

    # Data rows
    alt_fill = PatternFill("solid", fgColor="EBF3FB")
    for r, (hour, label) in enumerate(zip(HOURS, HOUR_LABELS), start=3):
        hc           = ws.cell(row=r, column=1, value=label)
        hc.font      = Font(bold=True, size=10)
        hc.alignment = Alignment(horizontal="center")
        hc.border    = _thin_border()
        if r % 2 == 0:
            hc.fill = alt_fill
        for c, day in enumerate(days, start=2):
            val            = pivot_data.get(day, {}).get(hour, 0)
            cell           = ws.cell(row=r, column=c, value=val)
            cell.alignment = Alignment(horizontal="center")
            cell.border    = _thin_border()
            if r % 2 == 0:
                cell.fill = alt_fill

    # Total row
    total_row  = 3 + len(HOURS)
    tot_fill   = PatternFill("solid", fgColor="1F4E79")
    tot_font   = Font(bold=True, color="FFFFFF", size=10)
    tc2           = ws.cell(row=total_row, column=1, value="Total")
    tc2.font      = tot_font
    tc2.fill      = tot_fill
    tc2.alignment = Alignment(horizontal="center")
    tc2.border    = _thin_border()
    for c, day in enumerate(days, start=2):
        day_total      = sum(pivot_data.get(day, {}).get(h, 0) for h in HOURS)
        cell           = ws.cell(row=total_row, column=c, value=day_total)
        cell.font      = tot_font
        cell.fill      = tot_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border    = _thin_border()

    # Column widths + freeze
    ws.column_dimensions["A"].width = 9
    for c in range(2, n_days + 2):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.freeze_panes = "B3"

    # ── Chart sheet ──────────────────────────────────────────────────────────
    wc       = wb.create_sheet("Chart")
    chart    = LineChart()
    chart.title        = title
    chart.y_axis.title = "Traffic Volume (vehicles)"
    chart.x_axis.title = "Hour of Day"
    chart.width        = 28
    chart.height       = 15
    chart.y_axis.numFmt = "0"
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    data_ref = Reference(ws, min_col=2, max_col=n_days + 1,
                         min_row=2, max_row=2 + len(HOURS))
    chart.add_data(data_ref, titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=3, max_row=2 + len(HOURS))
    chart.set_categories(cats)

    for i, series in enumerate(chart.series):
        color = DAY_COLORS[i % len(DAY_COLORS)]
        series.smooth = True
        series.graphicalProperties.line.solidFill          = color
        series.graphicalProperties.line.width              = 18000
        series.marker.symbol                               = "circle"
        series.marker.size                                 = 4
        series.marker.graphicalProperties.solidFill        = color
        series.marker.graphicalProperties.line.solidFill   = color

    legend          = Legend()
    legend.position = "r"
    legend.overlay  = False
    chart.legend    = legend

    wc.add_chart(chart, "B2")

    wb.save(out_path)
    print(f"  XLSX → {out_path}")


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for approach, directory in APPROACH_DIRS.items():
        print(f"\n{'─'*50}")
        print(f"Approach: {approach}  ({directory})")

        if not os.path.isdir(directory):
            print(f"  [SKIP] Directory not found: {directory}")
            continue

        zone_map = discover_zones(directory, approach)
        if not zone_map:
            print(f"  [SKIP] No matching zones found in {directory}")
            continue

        for zone, mappings in zone_map.items():
            desc = ", ".join(f"{mv}←{col}" for mv, col in mappings)
            print(f"  {zone:12s}  {desc}")

        all_pivot = load_all_files(directory, zone_map)

        for movement, pivot_data in all_pivot.items():
            days  = list(pivot_data.keys())
            label = MOVEMENT_LABELS.get(movement, movement)
            approach_label = APPROACH_LABELS.get(approach, approach)
            title = f"{approach_label} {label}\n{INTERSECTION}"
            col   = f"{approach}{movement}"
            stem  = f"{approach}_{movement}"

            csv_path  = os.path.join(OUTPUT_DIR, f"{stem}.csv")
            xlsx_path = os.path.join(OUTPUT_DIR, f"{stem}.xlsx")

            write_csv(pivot_data, days, col, csv_path)
            build_excel(pivot_data, days, title.replace("\n", " — "), xlsx_path)

    print(f"\nDone. All files written to '{OUTPUT_DIR}/'")


if __name__ == "__main__":
    main()
