import csv
from collections import defaultdict
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

INTERSECTION = "CSAH 61 (Flying Cloud Dr) at College View Dr"

GRAPHS = [
    {
        "csv":   "NB_THROUGH.csv",
        "col":   "NBT",
        "title": f"Northbound Through\n{INTERSECTION}",
        "out":   "NB_Through.png",
    },
    {
        "csv":   "NB_LEFT.csv",
        "col":   "NBL",
        "title": f"Northbound Left Turn\n{INTERSECTION}",
        "out":   "NB_Left.png",
    },
    {
        "csv":   "SB_THROUGH.csv",
        "col":   "SBT",
        "title": f"Southbound Through\n{INTERSECTION}",
        "out":   "SB_Through.png",
    },
    {
        "csv":   "SB_LEFT.csv",
        "col":   "SBL",
        "title": f"Southbound Left Turn\n{INTERSECTION}",
        "out":   "SB_Left.png",
    },
    {
        "csv":        "NB_THROUGH_DATA.csv",
        "col":        "NBT",
        "title":      f"Northbound Through\n{INTERSECTION}",
        "out":        "NBT_GRAPH.png",
        "legend_title": "Date (June 2026)",
    },
    {
        "csv":        "NB_LEFT_DATA.csv",
        "col":        "NBL",
        "title":      f"Northbound Left Turn\n{INTERSECTION}",
        "out":        "NBL_GRAPH.png",
        "legend_title": "Date (June 2026)",
    },
    {
        "csv":        "SB_THROUGH_DATA.csv",
        "col":        "SBT",
        "title":      f"Southbound Through\n{INTERSECTION}",
        "out":        "SBT_GRAPH.png",
        "legend_title": "Date (June 2026)",
    },
    {
        "csv":        "SB_LEFT_DATA.csv",
        "col":        "SBL",
        "title":      f"Southbound Left Turn\n{INTERSECTION}",
        "out":        "SBL_GRAPH.png",
        "legend_title": "Date (June 2026)",
    },
]

# 10 visually distinct colors for the 10 weekdays
DAY_COLORS = [
    "#e6194b", "#3cb44b", "#4363d8", "#f58231", "#911eb4",
    "#42d4f4", "#f032e6", "#bfef45", "#fabed4", "#469990",
]


def day_label(filename):
    """NB_MAY_4.csv -> 'May-4'"""
    stem = Path(filename).stem          # e.g. NB_MAY_4
    parts = stem.split("_")             # ['NB', 'MAY', '4']
    return f"{parts[-2].capitalize()}-{parts[-1]}"


def load_csv(csv_path, col):
    """Return {day_label: {hour: volume}} skipping the TOTAL row."""
    data = defaultdict(dict)
    with open(csv_path, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["File"] == "TOTAL":
                continue
            label = day_label(row["File"])
            hour  = int(row["Hour"])
            vol   = int(row[col])
            data[label][hour] = vol
    return data


def plot_graph(cfg):
    data = load_csv(cfg["csv"], cfg["col"])

    # Sort days chronologically by month-day number
    def sort_key(label):
        _, day = label.split("-")
        return int(day)

    days = sorted(data.keys(), key=sort_key)

    fig, ax = plt.subplots(figsize=(13, 6))
    fig.patch.set_facecolor("#f9f9f9")
    ax.set_facecolor("#ffffff")

    for i, day in enumerate(days):
        hours  = sorted(data[day].keys())
        values = [data[day][h] for h in hours]
        ax.plot(
            hours, values,
            color=DAY_COLORS[i],
            linewidth=1.8,
            marker="o",
            markersize=3.5,
            label=day,
        )

    # Axes labels & title
    ax.set_xlabel("Hour of Day", fontsize=12, labelpad=8)
    ax.set_ylabel("Traffic Volume (vehicles)", fontsize=12, labelpad=8)
    ax.set_title(cfg["title"], fontsize=14, fontweight="bold", pad=14)

    # X-axis: every hour 0-22
    all_hours = list(range(23))
    ax.set_xticks(all_hours)
    ax.set_xticklabels(
        [f"{h:02d}:00" for h in all_hours],
        rotation=45, ha="right", fontsize=8,
    )
    ax.set_xlim(-0.5, 22.5)

    # Y-axis: integer ticks, start at 0
    ax.set_ylim(bottom=0)
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True, nbins=8))

    # Grid
    ax.grid(axis="y", linestyle="--", linewidth=0.6, alpha=0.7, color="#cccccc")
    ax.grid(axis="x", linestyle=":",  linewidth=0.4, alpha=0.5, color="#dddddd")
    ax.set_axisbelow(True)

    # Legend
    ax.legend(
        title=cfg.get("legend_title", "Date (May 2026)"),
        title_fontsize=9,
        fontsize=8.5,
        loc="upper left",
        framealpha=0.85,
        edgecolor="#cccccc",
    )

    # Spine styling
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    for spine in ["left", "bottom"]:
        ax.spines[spine].set_color("#aaaaaa")

    plt.tight_layout()
    fig.savefig(cfg["out"], dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved {cfg['out']}")


if __name__ == "__main__":
    for cfg in GRAPHS:
        plot_graph(cfg)
