import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

INTERSECTION = "CSAH 61 (Flying Cloud Dr) at College View Dr"

EXPORTS = [
    {
        "csvs":  ["NB_THROUGH_DATA.csv"],
        "col":   "NBT",
        "title": f"Northbound Through — {INTERSECTION}",
        "out":   "NB_THROUGH_POINTS.xlsx",
    },
    {
        "csvs":  ["NB_LEFT_DATA.csv"],
        "col":   "NBL",
        "title": f"Northbound Left Turn — {INTERSECTION}",
        "out":   "NB_LEFT_POINTS.xlsx",
    },
    {
        "csvs":  ["SB_THROUGH_DATA.csv"],
        "col":   "SBT",
        "title": f"Southbound Through — {INTERSECTION}",
        "out":   "SB_THROUGH_POINTS.xlsx",
    },
    {
        "csvs":  ["SB_LEFT_DATA.csv"],
        "col":   "SBL",
        "title": f"Southbound Left Turn — {INTERSECTION}",
        "out":   "SB_LEFT_POINTS.xlsx",
    },
]

HOURS = list(range(23))


def day_label(filename):
    """NB_THROUGH.csv rows have File like 'NB_MAY_4.csv' or 'NB_June_1.csv'."""
    stem = Path(filename).stem          # e.g. NB_MAY_4 or NB_June_1
    parts = stem.split("_")
    return f"{parts[-2].capitalize()}-{parts[-1]}"


def load_pivoted(csv_files, col):
    """Return (ordered list of day labels, {day: {hour: volume}})."""
    day_data = {}
    day_order = []

    for csv_path in csv_files:
        with open(csv_path, newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row["File"] == "TOTAL":
                    continue
                label = day_label(row["File"])
                hour  = int(row["Hour"])
                vol   = int(row[col])
                if label not in day_data:
                    day_data[label] = {}
                    day_order.append(label)
                day_data[label][hour] = vol

    return day_order, day_data


def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def build_xlsx(cfg):
    days, data = load_pivoted(cfg["csvs"], cfg["col"])

    wb = Workbook()

    # ── Data sheet ──────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Data"

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(days) + 1)
    title_cell = ws.cell(row=1, column=1, value=cfg["title"])
    title_cell.font       = Font(bold=True, size=13)
    title_cell.alignment  = Alignment(horizontal="center", vertical="center")
    title_cell.fill       = PatternFill("solid", fgColor="1F4E79")
    title_cell.font       = Font(bold=True, size=13, color="FFFFFF")
    ws.row_dimensions[1].height = 24

    # Header row — Hour + day labels
    header_fill = PatternFill("solid", fgColor="2E75B6")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    ws.cell(row=2, column=1, value="Hour").font = Font(bold=True, color="FFFFFF", size=10)
    ws.cell(row=2, column=1).fill      = header_fill
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=2, column=1).border    = thin_border()

    for c, day in enumerate(days, start=2):
        cell = ws.cell(row=2, column=c, value=day)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border    = thin_border()

    # Data rows
    alt_fill = PatternFill("solid", fgColor="EBF3FB")
    for r, hour in enumerate(HOURS, start=3):
        hour_cell = ws.cell(row=r, column=1, value=hour)
        hour_cell.font      = Font(bold=True, size=10)
        hour_cell.alignment = Alignment(horizontal="center")
        hour_cell.border    = thin_border()
        if r % 2 == 0:
            hour_cell.fill = alt_fill

        for c, day in enumerate(days, start=2):
            val  = data[day].get(hour, 0)
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = Alignment(horizontal="center")
            cell.border    = thin_border()
            if r % 2 == 0:
                cell.fill = alt_fill

    # Column widths
    ws.column_dimensions["A"].width = 8
    for c in range(2, len(days) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 11

    ws.freeze_panes = "B3"

    # ── Chart sheet ─────────────────────────────────────────────────────────
    wc = wb.create_sheet("Chart")

    chart = LineChart()
    chart.title    = cfg["title"]
    chart.y_axis.title = "Traffic Volume (vehicles)"
    chart.x_axis.title = "Hour of Day"
    chart.style    = 10
    chart.width    = 24
    chart.height   = 14
    chart.y_axis.numFmt = "0"
    chart.y_axis.majorGridlines = None   # cleaner look; gridlines still show via style

    # All data columns (row 2 = headers, rows 3+ = values)
    data_ref = Reference(ws, min_col=2, max_col=len(days) + 1,
                         min_row=2, max_row=2 + len(HOURS))
    chart.add_data(data_ref, titles_from_data=True)

    # X-axis categories (hours)
    cats = Reference(ws, min_col=1, min_row=3, max_row=2 + len(HOURS))
    chart.set_categories(cats)

    # Smooth all series
    for series in chart.series:
        series.smooth = True

    wc.add_chart(chart, "B2")

    wb.save(cfg["out"])
    print(f"Saved {cfg['out']}")


if __name__ == "__main__":
    for cfg in EXPORTS:
        build_xlsx(cfg)
