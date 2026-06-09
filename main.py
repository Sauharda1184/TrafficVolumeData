import csv
import os
from collections import defaultdict

NB_DIR = "NorthBound"
NB_FILES = [
    "NB_MAY_4.csv", "NB_MAY_5.csv", "NB_MAY_6.csv", "NB_MAY_7.csv", "NB_MAY_8.csv",
    "NB_MAY_11.csv", "NB_MAY_12.csv", "NB_MAY_13.csv", "NB_MAY_14.csv", "NB_MAY_15.csv",
]

SB_DIR = "SouthBound"
SB_FILES = [
    "SB_MAY_4.csv", "SB_MAY_5.csv", "SB_MAY_6.csv", "SB_MAY_7.csv", "SB_MAY_8.csv",
    "SB_MAY_11.csv", "SB_MAY_12.csv", "SB_MAY_13.csv", "SB_MAY_14.csv", "SB_MAY_15.csv",
]


def process_file(filepath, through_zones, left_zone):
    """Return list of (hour, through, left) summed across all 15-min intervals in that hour."""
    hourly = defaultdict(lambda: {"T": 0, "L": 0})

    with open(filepath, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            zone = row["ZoneName"].strip()
            ts = row["TimeStamp"].strip()
            volume = int(float(row["Volume"]))
            hour = int(ts[11:13])

            if zone in through_zones:
                hourly[hour]["T"] += volume
            elif zone == left_zone:
                hourly[hour]["L"] += volume

    return [(hour, d["T"], d["L"]) for hour, d in sorted(hourly.items())]


def write_csvs(rows, through_file, left_file, through_col, left_col):
    total_t = sum(r["T"] for r in rows)
    total_l = sum(r["L"] for r in rows)

    with open(through_file, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["File", "Hour", through_col])
        writer.writeheader()
        for r in rows:
            writer.writerow({"File": r["File"], "Hour": r["Hour"], through_col: r["T"]})
        writer.writerow({"File": "TOTAL", "Hour": "", through_col: total_t})

    with open(left_file, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["File", "Hour", left_col])
        writer.writeheader()
        for r in rows:
            writer.writerow({"File": r["File"], "Hour": r["Hour"], left_col: r["L"]})
        writer.writerow({"File": "TOTAL", "Hour": "", left_col: total_l})

    print(f"Written {len(rows)} hourly rows to {through_file} and {left_file}")
    print(f"Grand totals — {through_col}: {total_t}, {left_col}: {total_l}")


def main():
    nb_rows = []
    for filename in NB_FILES:
        for hour, t, l in process_file(os.path.join(NB_DIR, filename), {"NBT1", "NBT2"}, "NBL1"):
            nb_rows.append({"File": filename, "Hour": hour, "T": t, "L": l})

    write_csvs(nb_rows, "NB_THROUGH.csv", "NB_LEFT.csv", "NBT", "NBL")

    sb_rows = []
    for filename in SB_FILES:
        for hour, t, l in process_file(os.path.join(SB_DIR, filename), {"SBT1", "SBT2"}, "SBL1"):
            sb_rows.append({"File": filename, "Hour": hour, "T": t, "L": l})

    write_csvs(sb_rows, "SB_THROUGH.csv", "SB_LEFT.csv", "SBT", "SBL")


if __name__ == "__main__":
    main()
