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
from datetime import date
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

# If True, write the standard per-approach/movement CSV + XLSX (with charts).
FULL_ANALYSIS_EXPORT = True

# If True, also write a single consolidated clean summary workbook combining
# every approach+movement into one chart-free, expandable table (Hour rows
# with collapsible 15-min detail, one column per approach+movement + Total).
CLEAN_SUMMARY_EXPORT = False

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

HOURS       = list(range(23))
HOUR_LABELS = [f"{h:02d}:00" for h in HOURS]
INTERVALS   = [f"{h:02d}:{m:02d}" for h in range(23) for m in (0, 15, 30, 45)]

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

# Scans every CSV in a folder and collects every distinct ZoneName found.
def scan_all_zones(directory):
    """Return sorted list of every unique ZoneName found across all CSVs in directory."""
    zones = set()
    for fname in sorted(Path(directory).glob("*.csv")):
        with open(fname, newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                z = row.get("ZoneName", "").strip()
                if z:
                    zones.add(z)
    return sorted(zones)


# Converts the GUI's per-zone column choices into the zone_map format the processing functions expect.
def config_to_zone_map(zone_config):
    """
    Convert GUI zone config to the zone_map format used by process_file_15min().

    zone_config: {zone_name: {"T": col_or_None, "L": col_or_None, "R": col_or_None}}
    Returns:     {zone_name: [(movement, column), ...]}   (only non-Skip entries)
    """
    zone_map = {}
    for zone, mv_cols in zone_config.items():
        mappings = [(mv, col) for mv, col in mv_cols.items() if col and col != "Skip"]
        if mappings:
            zone_map[zone] = mappings
    return zone_map


# Scans every CSV in a folder, finds all distinct ZoneName values, and pattern-matches them
# to guess which movement each zone feeds and which CSV column holds the counts.
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


# Converts a data filename into a short, readable day label for column headers.
def day_label(filename):
    """
    Convert a filename to a readable day label.
    E.g. 'NB_June_1.csv' -> 'June-1'
         'NB_MAY_4.csv'  -> 'May-4'
    """
    stem  = Path(filename).stem
    parts = stem.split("_")
    return f"{parts[-2].capitalize()}-{parts[-1]}"


# Reads one day's CSV and sums volumes into 15-minute buckets per movement,
# using whatever zone map was resolved by discover_zones() or a custom config.
def process_file_15min(filepath, zone_map):
    """
    Read one CSV file and return raw 15-minute interval totals per movement.
    zone_map: { zone_name: [(movement, column), ...] }

    Returns: {movement: {"HH:MM": total_volume}}
    """
    intervals = defaultdict(lambda: defaultdict(int))

    with open(filepath, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            zone = row.get("ZoneName", "").strip()
            if zone not in zone_map:
                continue
            ts       = row["TimeStamp"].strip()
            interval = ts[11:16]              # "HH:MM"
            for movement, column in zone_map[zone]:
                value = int(float(row.get(column, 0) or 0))
                intervals[movement][interval] += value

    return {mv: dict(ivs) for mv, ivs in intervals.items()}


# Runs process_file_15min() across every CSV in a folder (one per day, named like SB_June_1.csv).
def load_all_files_15min(directory, zone_map):
    """
    Process every CSV in *directory* and return pivoted 15-min data per movement.
    Returns: {movement: {day_label: {"HH:MM": volume}}}
    """
    pivot = defaultdict(dict)

    for fname in sorted(Path(directory).glob("*.csv"),
                        key=lambda p: [int(x) if x.isdigit() else x
                                       for x in re.split(r"(\d+)", p.stem)]):
        label     = day_label(fname.name)
        file_data = process_file_15min(fname, zone_map)
        for movement, iv_data in file_data.items():
            pivot[movement][label] = iv_data

    return dict(pivot)


# Derives hourly totals from a 15-minute pivot, so the hourly and 15-minute
# views always come from a single source of truth instead of two separate
# read-and-aggregate passes over the CSVs.
def hourly_from_15min(pivot_15min):
    """
    Sum each hour's four quarter-hour intervals to build the hourly pivot.
    pivot_15min: {movement: {day_label: {"HH:MM": volume}}}
    Returns:     {movement: {day_label: {hour: volume}}}
    """
    hourly = {}
    for movement, day_data in pivot_15min.items():
        hourly[movement] = {}
        for label, intervals in day_data.items():
            hours_seen = {int(iv[:2]) for iv in intervals}
            hourly[movement][label] = {
                hour: sum(intervals.get(f"{hour:02d}:{m:02d}", 0) for m in (0, 15, 30, 45))
                for hour in hours_seen
            }
    return hourly


# ─────────────────────────────────────────────────────────────────────────────
# CSV output
# ─────────────────────────────────────────────────────────────────────────────

# Writes the full-analysis hourly pivot table (one approach+movement, all days) to a plain CSV.
def write_csv(pivot_data, days, col_name, out_path):
    """Write hourly pivot table + TOTAL row to a CSV file."""
    n = len(days)
    with open(out_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Hour"] + days + ["Average"])
        for hour, label in zip(HOURS, HOUR_LABELS):
            vals = [pivot_data.get(d, {}).get(hour, 0) for d in days]
            avg  = round(sum(vals) / n, 1) if n else 0
            writer.writerow([label] + vals + [avg])
        totals = [sum(pivot_data.get(d, {}).get(h, 0) for h in HOURS) for d in days]
        avg_total = round(sum(totals) / n, 1) if n else 0
        writer.writerow(["Total"] + totals + [avg_total])
    print(f"  CSV  → {out_path}")


# Column order used within each approach in the clean summary export.
CLEAN_MOVEMENT_ORDER = ["L", "T", "R"]


# Builds the '<Intersection>_<Date>.xlsx' filename used for the clean summary export.
def clean_summary_filename(intersection, when=None):
    """
    Build a safe '<Intersection>_<YYYY-MM-DD>.xlsx' filename for the clean
    summary export, e.g. 'CSAH_61_Flying_Cloud_Dr_at_College_View_Dr_2026-07-08.xlsx'.
    Defaults to today's date (the day the export is run).
    """
    when = when or date.today()
    safe = re.sub(r"[^\w\-]+", "_", intersection).strip("_")
    return f"{safe}_{when.isoformat()}.xlsx"


# ─────────────────────────────────────────────────────────────────────────────
# Excel output
# ─────────────────────────────────────────────────────────────────────────────

# Shared style objects reused across every sheet-building function below, so a
# color or font only ever needs to change in one place.
CENTER        = Alignment(horizontal="center")
CENTER_MIDDLE = Alignment(horizontal="center", vertical="center")
CENTER_INDENT = Alignment(horizontal="center", indent=2)

FONT_TITLE        = Font(bold=True, size=13, color="FFFFFF")
FONT_HEADER       = Font(bold=True, color="FFFFFF", size=10)
FONT_HEADER_SMALL = Font(bold=True, color="FFFFFF", size=9)
FONT_TOTAL        = Font(bold=True, color="FFFFFF", size=10)
FONT_HOUR_LABEL   = Font(bold=True, size=10)
FONT_SUM_ROW      = Font(bold=True, size=10)
FONT_DET_ROW      = Font(size=9)
FONT_SUM_LAST     = Font(bold=True, size=10, color="202020")
FONT_DET_LAST     = Font(size=9, color="202020")
FONT_AVG_CELL     = Font(bold=True, size=10, color="202020")
FONT_INSTRUCTION  = Font(size=9, italic=True, color="444444")

FILL_TITLE       = PatternFill("solid", fgColor="1F4E79")
FILL_HEADER      = PatternFill("solid", fgColor="2E75B6")
FILL_TOTAL       = PatternFill("solid", fgColor="1F4E79")
FILL_AVG         = PatternFill("solid", fgColor="404040")
FILL_ALT_ROW     = PatternFill("solid", fgColor="EBF3FB")
FILL_AVG_CELL    = PatternFill("solid", fgColor="E8E8E8")
FILL_AVG_ALT     = PatternFill("solid", fgColor="D8D8D8")
FILL_INSTRUCTION = PatternFill("solid", fgColor="EBF3FB")
FILL_SUM_ROW     = PatternFill("solid", fgColor="D6E4F0")
FILL_DET_ROW_A   = PatternFill("solid", fgColor="FFFFFF")
FILL_DET_ROW_B   = PatternFill("solid", fgColor="F5F9FD")
FILL_SUM_LAST    = PatternFill("solid", fgColor="E0E0E0")
FILL_DET_LAST    = PatternFill("solid", fgColor="F0F0F0")


# Builds the thin light-grey border applied to every table cell in every sheet.
def _thin_border():
    """Return a thin light-grey border used on every data cell."""
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


# Writes and styles the merged dark-blue title banner that sits in row 1 of every sheet.
def _style_title_row(ws, end_col, text):
    """Write and style the merged title banner in row 1 of a worksheet."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    tc = ws.cell(row=1, column=1, value=text)
    tc.font, tc.fill, tc.alignment = FONT_TITLE, FILL_TITLE, CENTER_MIDDLE
    ws.row_dimensions[1].height = 24


# Applies the standard blue header-row look to a single cell.
def _style_header_cell(cell):
    """Apply the standard blue header-row styling to a single cell."""
    cell.font, cell.fill = FONT_HEADER, FILL_HEADER
    cell.alignment, cell.border = CENTER, _thin_border()


# Sets one uniform column width across a contiguous range of columns.
def _set_column_widths(ws, start_col, end_col, width):
    """Set the same column width for every column index in [start_col, end_col]."""
    for c in range(start_col, end_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


# Colors and styles a line chart's series: one colored line per day, plus a
# thick dashed grey Average line for whatever series comes after them.
def _style_line_chart_series(chart, n_colored, series_width=18000, marker_size=4, avg_width=28000):
    """
    Color the first n_colored series (one per day) using DAY_COLORS with round
    markers; style any remaining series (the Average line) as a thick dashed
    grey line with no markers. Shared by both the hourly and 15-min charts.
    """
    for i, series in enumerate(chart.series):
        series.smooth = True
        if i < n_colored:
            color = DAY_COLORS[i % len(DAY_COLORS)]
            series.graphicalProperties.line.solidFill        = color
            series.graphicalProperties.line.width             = series_width
            series.marker.symbol                              = "circle"
            series.marker.size                                = marker_size
            series.marker.graphicalProperties.solidFill       = color
            series.marker.graphicalProperties.line.solidFill  = color
        else:
            series.graphicalProperties.line.solidFill  = "404040"
            series.graphicalProperties.line.width      = avg_width
            series.graphicalProperties.line.dashStyle  = "dash"
            series.marker.symbol                       = "none"


# Builds the standard right-side legend shared by every line chart sheet.
def _new_legend():
    """Build the standard right-side chart legend used by every line chart."""
    legend          = Legend()
    legend.position = "r"
    legend.overlay  = False
    return legend


# Builds the full per-approach/movement workbook: a flat Data sheet, a Chart
# sheet, and (optionally) the expandable 15-min Data + 15-min Chart sheets.
def build_excel(pivot_data, days, title, out_path, pivot_15min=None):
    """
    Build a styled Excel workbook with a flat "Data" sheet (hourly pivot +
    Average column) and a "Chart" sheet (line chart of the same data). If
    pivot_15min is given, also appends "15-min Data" (expandable) and
    "15-min Chart" sheets. Writes the workbook to out_path.
    """
    wb = Workbook()

    # ── Data sheet ───────────────────────────────────────────────────────────
    ws       = wb.active
    ws.title = "Data"
    n_days   = len(days)
    avg_col  = n_days + 2

    _style_title_row(ws, avg_col, title)

    _style_header_cell(ws.cell(row=2, column=1, value="Hour"))
    for c, day in enumerate(days, start=2):
        _style_header_cell(ws.cell(row=2, column=c, value=day))
    ahdr = ws.cell(row=2, column=avg_col, value="Average")
    ahdr.font, ahdr.fill = FONT_TOTAL, FILL_AVG
    ahdr.alignment, ahdr.border = CENTER, _thin_border()

    # Data rows (alternating shading every other hour)
    for r, (hour, label) in enumerate(zip(HOURS, HOUR_LABELS), start=3):
        alt = (r % 2 == 0)
        hc = ws.cell(row=r, column=1, value=label)
        hc.font, hc.alignment, hc.border = FONT_HOUR_LABEL, CENTER, _thin_border()
        if alt:
            hc.fill = FILL_ALT_ROW
        vals = []
        for c, day in enumerate(days, start=2):
            val = pivot_data.get(day, {}).get(hour, 0)
            vals.append(val)
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment, cell.border = CENTER, _thin_border()
            if alt:
                cell.fill = FILL_ALT_ROW
        avg_val = round(sum(vals) / n_days, 1) if n_days else 0
        ac = ws.cell(row=r, column=avg_col, value=avg_val)
        ac.font, ac.fill = FONT_AVG_CELL, (FILL_AVG_ALT if alt else FILL_AVG_CELL)
        ac.alignment, ac.border = CENTER, _thin_border()

    # Total row
    total_row = 3 + len(HOURS)
    tc = ws.cell(row=total_row, column=1, value="Total")
    tc.font, tc.fill, tc.alignment, tc.border = FONT_TOTAL, FILL_TOTAL, CENTER, _thin_border()
    day_totals = []
    for c, day in enumerate(days, start=2):
        day_total = sum(pivot_data.get(day, {}).get(h, 0) for h in HOURS)
        day_totals.append(day_total)
        cell = ws.cell(row=total_row, column=c, value=day_total)
        cell.font, cell.fill, cell.alignment, cell.border = FONT_TOTAL, FILL_TOTAL, CENTER, _thin_border()
    avg_total = round(sum(day_totals) / n_days, 1) if n_days else 0
    atc = ws.cell(row=total_row, column=avg_col, value=avg_total)
    atc.font, atc.fill, atc.alignment, atc.border = FONT_TOTAL, FILL_AVG, CENTER, _thin_border()

    ws.column_dimensions["A"].width = 9
    _set_column_widths(ws, 2, avg_col, 11)
    ws.freeze_panes = "B3"

    # ── Chart sheet ──────────────────────────────────────────────────────────
    wc    = wb.create_sheet("Chart")
    chart = LineChart()
    chart.title, chart.y_axis.title, chart.x_axis.title = title, "Traffic Volume (vehicles)", "Hour of Day"
    chart.width, chart.height = 28, 15
    chart.y_axis.numFmt  = "0"
    chart.x_axis.delete  = chart.y_axis.delete = False

    data_ref = Reference(ws, min_col=2, max_col=avg_col, min_row=2, max_row=2 + len(HOURS))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=3, max_row=2 + len(HOURS)))
    _style_line_chart_series(chart, n_days)
    chart.legend = _new_legend()
    wc.add_chart(chart, "B2")

    # ── 15-min sheets (optional) ─────────────────────────────────────────────
    if pivot_15min is not None:
        _build_expandable_hour_sheet(
            wb, "15-min Data", f"{title} — 15-Minute Intervals", days,
            value_lookup=lambda day, iv: pivot_15min.get(day, {}).get(iv, 0),
            hours=HOURS,
            last_col_label="Average",
            last_col_fn=lambda vals: round(sum(vals) / n_days, 1) if n_days else 0,
        )
        _build_15min_chart_sheet(wb, pivot_15min, days, title, n_days)

    wb.save(out_path)
    print(f"  XLSX → {out_path}")


# Builds one worksheet where each hour is a collapsible group: a bold summary
# row plus four hidden 15-minute detail rows beneath it. Shared by the
# per-movement "15-min Data" sheet (columns=days, last column=Average) and the
# Clean Summary workbook (columns=approach+movement codes, last column=Total).
def _build_expandable_hour_sheet(wb, sheet_name, title, columns, value_lookup,
                                  hours, last_col_label, last_col_fn):
    """
    Build a worksheet where each hour is a bold summary row followed by four
    collapsed 15-minute detail rows (click [+] to expand) — one column per
    entry in `columns`, plus a final aggregate column, and a Total row at the
    bottom.

    value_lookup(column, "HH:MM") -> int          volume for one column/interval
    hours                                          which hours (0-22) to render as rows
    last_col_fn(row_values: list[int]) -> number   value for the aggregate column
    Returns the created worksheet.
    """
    ws       = wb.create_sheet(sheet_name)
    last_col = len(columns) + 2

    # Summary rows appear ABOVE their detail rows so [+] sits on the summary row
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False

    _style_title_row(ws, last_col, title)

    # Instruction row
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ic = ws.cell(row=2, column=1,
                 value="Click [+] on the left margin to expand an hour into its four 15-minute intervals")
    ic.font, ic.fill, ic.alignment = FONT_INSTRUCTION, FILL_INSTRUCTION, CENTER
    ws.row_dimensions[2].height = 16

    # Header row
    _style_header_cell(ws.cell(row=3, column=1, value="Hour / Interval"))
    for c, col in enumerate(columns, start=2):
        _style_header_cell(ws.cell(row=3, column=c, value=col))
    lhdr = ws.cell(row=3, column=last_col, value=last_col_label)
    lhdr.font, lhdr.fill = FONT_TOTAL, FILL_AVG
    lhdr.alignment, lhdr.border = CENTER, _thin_border()

    r = 4
    col_totals = defaultdict(int)
    for hour in hours:
        # ── Summary row (hourly total across the 4 quarters) ─────────────────
        sc = ws.cell(row=r, column=1, value=f"{hour:02d}:00")
        sc.font, sc.fill = FONT_SUM_ROW, FILL_SUM_ROW
        sc.alignment, sc.border = CENTER, _thin_border()
        row_vals = []
        for c, col in enumerate(columns, start=2):
            val = sum(value_lookup(col, f"{hour:02d}:{m:02d}") for m in (0, 15, 30, 45))
            row_vals.append(val)
            cell = ws.cell(row=r, column=c, value=val)
            cell.font, cell.fill = FONT_SUM_ROW, FILL_SUM_ROW
            cell.alignment, cell.border = CENTER, _thin_border()
        for col, v in zip(columns, row_vals):
            col_totals[col] += v
        slc = ws.cell(row=r, column=last_col, value=last_col_fn(row_vals))
        slc.font, slc.fill = FONT_SUM_LAST, FILL_SUM_LAST
        slc.alignment, slc.border = CENTER, _thin_border()
        r += 1

        # ── Detail rows (collapsed by default) ───────────────────────────────
        for i, m in enumerate((0, 15, 30, 45)):
            interval = f"{hour:02d}:{m:02d}"
            fill = FILL_DET_ROW_B if i % 2 else FILL_DET_ROW_A
            dc = ws.cell(row=r, column=1, value=interval)
            dc.font, dc.fill = FONT_DET_ROW, fill
            dc.alignment, dc.border = CENTER_INDENT, _thin_border()
            iv_vals = []
            for c, col in enumerate(columns, start=2):
                val = value_lookup(col, interval)
                iv_vals.append(val)
                cell = ws.cell(row=r, column=c, value=val)
                cell.font, cell.fill = FONT_DET_ROW, fill
                cell.alignment, cell.border = CENTER, _thin_border()
            dlc = ws.cell(row=r, column=last_col, value=last_col_fn(iv_vals))
            dlc.font, dlc.fill = FONT_DET_LAST, FILL_DET_LAST
            dlc.alignment, dlc.border = CENTER, _thin_border()
            ws.row_dimensions[r].outline_level = 1
            ws.row_dimensions[r].hidden        = True
            r += 1

    # Total row
    tc = ws.cell(row=r, column=1, value="Total")
    tc.font, tc.fill, tc.alignment, tc.border = FONT_TOTAL, FILL_TOTAL, CENTER, _thin_border()
    totals_row = [col_totals[c] for c in columns]
    for c, val in zip(range(2, last_col), totals_row):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font, cell.fill, cell.alignment, cell.border = FONT_TOTAL, FILL_TOTAL, CENTER, _thin_border()
    gc = ws.cell(row=r, column=last_col, value=last_col_fn(totals_row))
    gc.font, gc.fill, gc.alignment, gc.border = FONT_TOTAL, FILL_AVG, CENTER, _thin_border()

    ws.column_dimensions["A"].width = 15
    _set_column_widths(ws, 2, last_col, 11)
    ws.freeze_panes = "B4"
    return ws


# Builds the flat 92-row 15-minute data table plus its own line chart, on a dedicated sheet.
def _build_15min_chart_sheet(wb, pivot_15min, days, title, n_days):
    """Build a flat 92-row 15-minute data table plus a line chart, on a dedicated sheet."""
    wc      = wb.create_sheet("15-min Chart")
    avg_col = n_days + 2

    h = wc.cell(row=1, column=1, value="Interval")
    h.font, h.fill, h.alignment = FONT_HEADER_SMALL, FILL_HEADER, CENTER
    for c, day in enumerate(days, start=2):
        cell = wc.cell(row=1, column=c, value=day)
        cell.font, cell.fill, cell.alignment = FONT_HEADER_SMALL, FILL_HEADER, CENTER
    ahdr = wc.cell(row=1, column=avg_col, value="Average")
    ahdr.font, ahdr.fill, ahdr.alignment = FONT_HEADER_SMALL, FILL_AVG, CENTER

    for r, interval in enumerate(INTERVALS, start=2):
        wc.cell(row=r, column=1, value=interval)
        iv_vals = []
        for c, day in enumerate(days, start=2):
            val = pivot_15min.get(day, {}).get(interval, 0)
            iv_vals.append(val)
            wc.cell(row=r, column=c, value=val)
        avg_val = round(sum(iv_vals) / n_days, 1) if n_days else 0
        wc.cell(row=r, column=avg_col, value=avg_val)

    chart = LineChart()
    chart.title = f"{title} — 15-Minute Intervals"
    chart.y_axis.title, chart.x_axis.title = "Traffic Volume (vehicles)", "15-Minute Interval"
    chart.width, chart.height = 32, 16
    chart.y_axis.numFmt = "0"
    chart.x_axis.delete = chart.y_axis.delete = False

    data_ref = Reference(wc, min_col=2, max_col=avg_col, min_row=1, max_row=1 + len(INTERVALS))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(Reference(wc, min_col=1, min_row=2, max_row=1 + len(INTERVALS)))
    _style_line_chart_series(chart, n_days, series_width=15000, marker_size=3, avg_width=25000)
    chart.legend = _new_legend()
    wc.add_chart(chart, f"{get_column_letter(avg_col + 2)}1")

    wc.column_dimensions["A"].width = 9
    _set_column_widths(wc, 2, avg_col, 10)


# ─────────────────────────────────────────────────────────────────────────────
# Clean summary output (single day, all approaches, no charts)
# ─────────────────────────────────────────────────────────────────────────────

# Fixed approach column order for the clean summary, regardless of the order
# approaches were browsed/configured in: Hour, EBL EBT EBR WBL WBT WBR NBL NBT NBR SBL SBT SBR, Total
CLEAN_APPROACH_ORDER = ["EB", "WB", "NB", "SB"]


# Builds the single consolidated, chart-free, expandable workbook combining every approach+movement.
def build_clean_summary_excel(approach_pivots_15, approach_order, intersection, out_path,
                               movement_order=CLEAN_MOVEMENT_ORDER):
    """
    Build a single-sheet, chart-free Excel workbook for one day of data across
    every approach: one summary row per hour actually present in the source
    data, with four collapsible 15-minute detail rows beneath it (click [+] to
    expand), one column per approach+movement (always ordered EB, WB, NB, SB
    regardless of caller order — e.g. EBL, EBT, EBR, WBL, ...), plus a Total
    column and a Total row. Meant for quick comparison against
    consultant-provided counts.

    approach_pivots_15: {approach: {movement: {day_label: {"HH:MM": volume}}}}
                         — the dict returned by load_all_files_15min() for each approach.
    approach_order:     approaches actually present in this run; any not in
                         CLEAN_APPROACH_ORDER are appended at the end in the
                         order given.
    """
    ordered_approaches = sorted(
        approach_order,
        key=lambda a: (CLEAN_APPROACH_ORDER.index(a)
                        if a in CLEAN_APPROACH_ORDER else len(CLEAN_APPROACH_ORDER))
    )

    columns = [
        f"{a}{m}"
        for a in ordered_approaches
        for m in movement_order
        if m in approach_pivots_15.get(a, {})
    ]

    # Sum every day/file found for each approach+movement into one combined interval map.
    col_interval = defaultdict(lambda: defaultdict(int))
    for a in ordered_approaches:
        for m in movement_order:
            mv_data = approach_pivots_15.get(a, {}).get(m)
            if not mv_data:
                continue
            col = f"{a}{m}"
            for iv_data in mv_data.values():
                for interval, val in iv_data.items():
                    col_interval[col][interval] += val

    hours_present = sorted({int(iv[:2]) for ivs in col_interval.values() for iv in ivs})

    wb = Workbook()
    wb.remove(wb.active)   # the shared sheet builder creates its own named sheet
    _build_expandable_hour_sheet(
        wb, "Clean Summary", f"{intersection} — Clean Summary", columns,
        value_lookup=lambda col, iv: col_interval[col].get(iv, 0),
        hours=hours_present,
        last_col_label="Total",
        last_col_fn=sum,
    )
    wb.save(out_path)
    print(f"  Clean XLSX → {out_path}")


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline orchestration — shared by the CLI entry point (main()) and gui.py
# ─────────────────────────────────────────────────────────────────────────────

# Default logger used when no log callback is supplied (plain stdout printing, ignores tags).
def _console_log(msg, tag=""):
    """Print a pipeline log message to the console, ignoring the GUI-only tag."""
    print(msg)


# Runs the full pipeline for every approach: resolve zones, load data, and
# write whichever outputs are requested. Shared by main() (CLI) and gui.py
# (GUI) so both stay in sync through one implementation instead of two.
def run_pipeline(intersection, approach_dirs, output_dir, zone_configs=None,
                  full_analysis=True, clean_summary=False, log=_console_log):
    """
    Run the traffic-volume pipeline across every approach in approach_dirs.

    approach_dirs: {approach_code: directory_path}
    zone_configs:  optional {approach_code: {zone: {movement: column}}} — a
                   custom mapping (e.g. from the GUI's Configure Zones dialog);
                   any approach missing from this dict falls back to auto-detect.
    log:           callable(message, tag="") used for progress output — the
                   GUI passes its own log-widget writer to get colored tags.
    Returns the clean summary file path if one was written, else None.
    """
    zone_configs = zone_configs or {}
    os.makedirs(output_dir, exist_ok=True)
    log(f"Intersection : {intersection}", "info")
    log(f"Output folder: {output_dir}", "dim")
    log("─" * 56, "dim")

    approach_pivots_15 = {}   # {approach: {movement: {day_label: {"HH:MM": volume}}}}

    for approach, directory in approach_dirs.items():
        log(f"\n▸ {approach} ({APPROACH_LABELS.get(approach, approach)})  ← {directory}", "info")

        if not os.path.isdir(directory):
            log("  [SKIP] Directory not found.", "err")
            continue

        if approach in zone_configs:
            zone_map = config_to_zone_map(zone_configs[approach])
            log("  Using custom zone configuration:", "dim")
        else:
            zone_map = discover_zones(directory, approach)
            log("  Using auto-detected zone configuration:", "dim")

        if not zone_map:
            log("  [SKIP] No active zone assignments found.", "err")
            continue

        for zone, mappings in zone_map.items():
            desc = "  +  ".join(f"{mv}←{col}" for mv, col in mappings)
            log(f"  {zone:14s}  {desc}", "dim")

        all_pivot_15 = load_all_files_15min(directory, zone_map)
        all_pivot    = hourly_from_15min(all_pivot_15)
        approach_pivots_15[approach] = all_pivot_15

        if full_analysis:
            for movement, pivot_data in all_pivot.items():
                days  = list(pivot_data.keys())
                label = MOVEMENT_LABELS.get(movement, movement)
                approach_label = APPROACH_LABELS.get(approach, approach)
                title = f"{approach_label} {label} — {intersection}"
                stem  = f"{approach}_{movement}"

                csv_path  = os.path.join(output_dir, f"{stem}.csv")
                xlsx_path = os.path.join(output_dir, f"{stem}.xlsx")

                write_csv(pivot_data, days, f"{approach}{movement}", csv_path)
                log(f"  CSV  → {os.path.basename(csv_path)}", "ok")
                build_excel(pivot_data, days, title, xlsx_path,
                            pivot_15min=all_pivot_15.get(movement))
                log(f"  XLSX → {os.path.basename(xlsx_path)}", "ok")

    summary_path = None
    if clean_summary:
        summary_path = os.path.join(output_dir, clean_summary_filename(intersection))
        build_clean_summary_excel(approach_pivots_15, list(approach_dirs.keys()),
                                   intersection, summary_path)
        log(f"  Clean XLSX → {os.path.basename(summary_path)}", "ok")

    log("\n" + "─" * 56, "dim")
    log("✓  Done. All files written to output folder.", "ok")
    return summary_path


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

# CLI entry point: runs the pipeline using the hardcoded CONFIG section above.
def main():
    """Run the pipeline once using the CONFIG constants defined at the top of this file."""
    run_pipeline(
        INTERSECTION, APPROACH_DIRS, OUTPUT_DIR,
        full_analysis=FULL_ANALYSIS_EXPORT,
        clean_summary=CLEAN_SUMMARY_EXPORT,
    )


if __name__ == "__main__":
    main()
