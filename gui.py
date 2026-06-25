"""
gui.py
Tkinter GUI for the Traffic Volume Analysis pipeline.
Run:  python3 gui.py
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import volume_data as vd

APPROACHES = ["NB", "SB", "EB", "WB"]
APPROACH_FULL = {
    "NB": "Northbound",
    "SB": "Southbound",
    "EB": "Eastbound",
    "WB": "Westbound",
}

COLUMNS     = ["Skip", "Volume", "ThroughCount", "LeftTurnCount", "RightTurnCount"]
MOVEMENTS   = ["T", "L", "R"]
MV_LABELS   = {"T": "Through", "L": "Left", "R": "Right"}

BG         = "#f0f4f8"
HEADER_BG  = "#1F4E79"
HEADER_FG  = "#ffffff"
ACCENT     = "#2E75B6"
BTN_FG     = "#ffffff"
LOG_BG     = "#1e1e2e"
LOG_FG     = "#cdd6f4"
LOG_OK     = "#a6e3a1"
LOG_ERR    = "#f38ba8"
LOG_INFO   = "#89dceb"


# ── Zone Configuration Dialog ─────────────────────────────────────────────────

class ZoneConfigDialog(tk.Toplevel):
    """
    Modal dialog that shows every zone found in a directory and lets the user
    assign which column feeds each movement (T / L / R) for that zone.
    Multiple zones can feed the same movement — their values are summed.

    Example for WB approach:
      WBT1   Through→Volume        Left→Skip          Right→Skip
      WBTR1  Through→ThroughCount  Left→Skip          Right→RightTurnCount
      WBL1   Through→Skip          Left→Volume        Right→Skip
    """

    def __init__(self, parent, approach, directory, existing_config):
        super().__init__(parent)
        self.title(f"Configure Zones — {approach} ({APPROACH_FULL.get(approach, approach)})")
        self.resizable(True, True)
        self.configure(bg=BG)
        self.grab_set()

        self.result      = None        # set to config dict on OK
        self._approach   = approach
        self._directory  = directory
        self._row_vars   = {}          # zone → {mv: StringVar}

        self._build(existing_config)
        self._center(parent)

    # ── Build UI ─────────────────────────────────────────────────────────────

    def _build(self, existing_config):
        # Header
        hdr = tk.Frame(self, bg=HEADER_BG)
        hdr.pack(fill="x")
        tk.Label(
            hdr,
            text=f"{self._approach} — Zone Column Assignment",
            bg=HEADER_BG, fg=HEADER_FG,
            font=("Helvetica", 12, "bold"), pady=10,
        ).pack()
        tk.Label(
            hdr,
            text="For each zone choose which column feeds Through, Left, and Right counts.\n"
                 "Multiple zones contributing to the same movement are summed together.",
            bg=HEADER_BG, fg="#a8c8e8",
            font=("Helvetica", 8), pady=4,
        ).pack()

        # Scrollable table
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True, padx=16, pady=12)

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        vsb    = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        table = tk.Frame(canvas, bg=BG)
        canvas.create_window((0, 0), window=table, anchor="nw")
        table.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # Column headers
        hdr_font = ("Helvetica", 9, "bold")
        for col, text in enumerate(["Zone", "Through Column", "Left Column", "Right Column"]):
            tk.Label(
                table, text=text,
                bg=ACCENT, fg=BTN_FG,
                font=hdr_font,
                padx=10, pady=5,
                relief="flat", width=18 if col > 0 else 16,
            ).grid(row=0, column=col, padx=1, pady=1, sticky="ew")

        # Scan all zones in directory
        all_zones    = vd.scan_all_zones(self._directory)
        auto_map     = vd.discover_zones(self._directory, self._approach)
        auto_config  = _zone_map_to_config(auto_map, all_zones)

        # Use existing config if provided, otherwise use auto-detected
        config = existing_config if existing_config else auto_config

        alt_bg = "#eaf2fb"
        for r, zone in enumerate(all_zones, start=1):
            bg = alt_bg if r % 2 == 0 else "white"
            tk.Label(
                table, text=zone,
                bg=bg, font=("Courier", 9, "bold"),
                padx=8, pady=4, anchor="w",
            ).grid(row=r, column=0, padx=1, pady=1, sticky="ew")

            self._row_vars[zone] = {}
            for c, mv in enumerate(MOVEMENTS, start=1):
                var = tk.StringVar(value=config.get(zone, {}).get(mv, "Skip"))
                self._row_vars[zone][mv] = var
                cb = ttk.Combobox(
                    table,
                    textvariable=var,
                    values=COLUMNS,
                    state="readonly",
                    width=16,
                )
                cb.grid(row=r, column=c, padx=4, pady=2)

        # Buttons
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(fill="x", padx=16, pady=(0, 14))

        tk.Button(
            btn_frame, text="Reset to Auto-Detect",
            command=lambda: self._reset(auto_config),
            bg="#888", fg=BTN_FG,
            font=("Helvetica", 9), relief="flat", padx=10, pady=5, cursor="hand2",
        ).pack(side="left")

        tk.Button(
            btn_frame, text="Cancel",
            command=self.destroy,
            bg="#c0392b", fg=BTN_FG,
            font=("Helvetica", 9, "bold"), relief="flat", padx=10, pady=5, cursor="hand2",
        ).pack(side="right", padx=(6, 0))

        tk.Button(
            btn_frame, text="OK — Save Configuration",
            command=self._ok,
            bg="#217346", fg=BTN_FG,
            font=("Helvetica", 9, "bold"), relief="flat", padx=10, pady=5, cursor="hand2",
        ).pack(side="right")

        # Reasonable window size
        self.update_idletasks()
        h = min(80 + len(all_zones) * 34, 600)
        self.geometry(f"720x{h}")

    def _reset(self, auto_config):
        for zone, mv_vars in self._row_vars.items():
            for mv, var in mv_vars.items():
                var.set(auto_config.get(zone, {}).get(mv, "Skip"))

    def _ok(self):
        self.result = {
            zone: {mv: var.get() for mv, var in mv_vars.items()}
            for zone, mv_vars in self._row_vars.items()
        }
        self.destroy()

    def _center(self, parent):
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


def _zone_map_to_config(zone_map, all_zones):
    """Convert zone_map {zone: [(mv, col)]} to config {zone: {mv: col}} for all zones."""
    config = {zone: {mv: "Skip" for mv in MOVEMENTS} for zone in all_zones}
    for zone, mappings in zone_map.items():
        if zone in config:
            for mv, col in mappings:
                config[zone][mv] = col
    return config


# ── Main Application ──────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Traffic Volume Analyzer")
        self.resizable(False, False)
        self.configure(bg=BG)
        # Stores zone config per approach: {approach: {zone: {mv: col}}}
        self._zone_configs = {}
        self._build_ui()
        self._center()

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=HEADER_BG)
        hdr.pack(fill="x")
        tk.Label(
            hdr,
            text="Traffic Volume Analyzer",
            bg=HEADER_BG, fg=HEADER_FG,
            font=("Helvetica", 16, "bold"), pady=12,
        ).pack()
        tk.Label(
            hdr,
            text="Hourly Volume Analysis  •  Excel & CSV Output",
            bg=HEADER_BG, fg="#a8c8e8",
            font=("Helvetica", 9), pady=2,
        ).pack()

        body = tk.Frame(self, bg=BG, padx=20, pady=16)
        body.pack(fill="both")

        # Intersection name
        self._section(body, "Intersection Name")
        self.intersection_var = tk.StringVar()
        tk.Entry(
            body, textvariable=self.intersection_var,
            font=("Helvetica", 10), width=58,
            relief="solid", bd=1,
        ).pack(anchor="w", pady=(0, 12))

        # Approach directories
        self._section(body, "Approach Data Directories")
        tk.Label(
            body,
            text="Browse to the folder containing CSV files for each approach. "
                 "Then click Configure Zones to review or override the column assignments.",
            bg=BG, fg="#555", font=("Helvetica", 8, "italic"),
            wraplength=560, justify="left",
        ).pack(anchor="w", pady=(0, 8))

        self.dir_vars     = {}
        self._cfg_buttons = {}
        self._cfg_labels  = {}
        for approach in APPROACHES:
            self._approach_row(body, approach)

        # Output directory
        self._section(body, "Output Folder")
        out_row = tk.Frame(body, bg=BG)
        out_row.pack(fill="x", pady=(0, 12))
        self.output_var = tk.StringVar(value=os.path.join(os.getcwd(), "output"))
        tk.Entry(
            out_row, textvariable=self.output_var,
            font=("Helvetica", 10), width=46,
            relief="solid", bd=1,
        ).pack(side="left")
        tk.Button(
            out_row, text="Browse",
            command=self._browse_output,
            bg=ACCENT, fg=BTN_FG,
            font=("Helvetica", 9, "bold"),
            relief="flat", padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=(6, 0))

        # Run button
        self.run_btn = tk.Button(
            body, text="▶  Run Analysis",
            command=self._run,
            bg="#217346", fg=BTN_FG,
            font=("Helvetica", 11, "bold"),
            relief="flat", padx=16, pady=8, cursor="hand2",
        )
        self.run_btn.pack(pady=(4, 14))

        # Log
        self._section(body, "Output Log")
        log_frame = tk.Frame(body, bg=LOG_BG, bd=1, relief="solid")
        log_frame.pack(fill="both")
        self.log = tk.Text(
            log_frame, height=12, width=68,
            bg=LOG_BG, fg=LOG_FG,
            font=("Courier", 9),
            state="disabled", relief="flat",
            padx=8, pady=6,
        )
        sb = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.log.tag_config("ok",   foreground=LOG_OK)
        self.log.tag_config("err",  foreground=LOG_ERR)
        self.log.tag_config("info", foreground=LOG_INFO)
        self.log.tag_config("dim",  foreground="#6c7086")

    def _section(self, parent, text):
        tk.Label(
            parent, text=text.upper(),
            bg=BG, fg=ACCENT,
            font=("Helvetica", 8, "bold"),
        ).pack(anchor="w", pady=(8, 2))
        tk.Frame(parent, bg=ACCENT, height=1).pack(fill="x", pady=(0, 6))

    def _approach_row(self, parent, approach):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=3)

        tk.Label(
            row,
            text=f"{approach}  ({APPROACH_FULL[approach]})",
            bg=BG, fg="#333",
            font=("Helvetica", 10), width=18, anchor="w",
        ).pack(side="left")

        var = tk.StringVar()
        self.dir_vars[approach] = var
        tk.Entry(
            row, textvariable=var,
            font=("Helvetica", 9), width=28,
            relief="solid", bd=1,
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            row, text="Browse",
            command=lambda a=approach: self._browse_approach(a),
            bg=ACCENT, fg=BTN_FG,
            font=("Helvetica", 9, "bold"),
            relief="flat", padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=(0, 6))

        cfg_btn = tk.Button(
            row, text="Configure Zones",
            command=lambda a=approach: self._open_zone_config(a),
            bg="#888", fg=BTN_FG,
            font=("Helvetica", 9),
            relief="flat", padx=8, pady=2,
            state="disabled", cursor="hand2",
        )
        cfg_btn.pack(side="left", padx=(0, 6))
        self._cfg_buttons[approach] = cfg_btn

        lbl = tk.Label(row, text="", bg=BG, fg="#555", font=("Helvetica", 8, "italic"))
        lbl.pack(side="left")
        self._cfg_labels[approach] = lbl

    # ── Browsing ─────────────────────────────────────────────────────────────

    def _browse_approach(self, approach):
        path = filedialog.askdirectory(
            title=f"Select {approach} ({APPROACH_FULL[approach]}) data folder"
        )
        if not path:
            return
        self.dir_vars[approach].set(path)
        # Clear any saved config so next Configure opens fresh auto-detect
        self._zone_configs.pop(approach, None)
        # Activate Configure Zones button and update status label
        self._cfg_buttons[approach].configure(state="normal", bg=ACCENT)
        self._update_cfg_label(approach)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_var.set(path)

    def _open_zone_config(self, approach):
        directory = self.dir_vars[approach].get().strip()
        if not directory or not os.path.isdir(directory):
            messagebox.showwarning("Invalid Directory",
                                   f"Browse to a valid folder for {approach} first.")
            return
        existing = self._zone_configs.get(approach)
        dlg = ZoneConfigDialog(self, approach, directory, existing)
        self.wait_window(dlg)
        if dlg.result is not None:
            self._zone_configs[approach] = dlg.result
            self._update_cfg_label(approach)

    def _update_cfg_label(self, approach):
        if approach in self._zone_configs:
            cfg   = self._zone_configs[approach]
            active = sum(
                1 for mv_cols in cfg.values()
                for col in mv_cols.values() if col != "Skip"
            )
            self._cfg_labels[approach].configure(
                text=f"Custom ({active} assignments)", fg="#217346"
            )
        else:
            self._cfg_labels[approach].configure(
                text="Auto-detect on run", fg="#888"
            )

    # ── Logging ──────────────────────────────────────────────────────────────

    def _log(self, msg, tag=""):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── Run ──────────────────────────────────────────────────────────────────

    def _run(self):
        intersection = self.intersection_var.get().strip()
        if not intersection:
            messagebox.showwarning("Missing Input", "Please enter an intersection name.")
            return
        approach_dirs = {
            a: v.get().strip()
            for a, v in self.dir_vars.items()
            if v.get().strip()
        }
        if not approach_dirs:
            messagebox.showwarning("Missing Input",
                                   "Select at least one approach directory.")
            return
        output_dir = self.output_var.get().strip()
        self.run_btn.configure(state="disabled", text="Running…")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        threading.Thread(
            target=self._run_analysis,
            args=(intersection, approach_dirs, output_dir),
            daemon=True,
        ).start()

    def _run_analysis(self, intersection, approach_dirs, output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            self._log(f"Intersection : {intersection}", "info")
            self._log(f"Output folder: {output_dir}", "dim")
            self._log("─" * 56, "dim")

            for approach, directory in approach_dirs.items():
                self._log(
                    f"\n▸ {approach} ({vd.APPROACH_LABELS.get(approach, approach)})"
                    f"  ← {directory}", "info"
                )
                if not os.path.isdir(directory):
                    self._log("  [SKIP] Directory not found.", "err")
                    continue

                # Use saved config or auto-detect
                if approach in self._zone_configs:
                    zone_map = vd.config_to_zone_map(self._zone_configs[approach])
                    self._log("  Using custom zone configuration:", "dim")
                else:
                    zone_map = vd.discover_zones(directory, approach)
                    self._log("  Using auto-detected zone configuration:", "dim")

                if not zone_map:
                    self._log("  [SKIP] No active zone assignments found.", "err")
                    continue

                for zone, mappings in zone_map.items():
                    desc = "  +  ".join(f"{mv}←{col}" for mv, col in mappings)
                    self._log(f"  {zone:14s}  {desc}", "dim")

                all_pivot    = vd.load_all_files(directory, zone_map)
                all_pivot_15 = vd.load_all_files_15min(directory, zone_map)

                for movement, pivot_data in all_pivot.items():
                    days      = list(pivot_data.keys())
                    mv_label  = vd.MOVEMENT_LABELS.get(movement, movement)
                    a_label   = vd.APPROACH_LABELS.get(approach, approach)
                    title     = f"{a_label} {mv_label} — {intersection}"
                    stem      = f"{approach}_{movement}"
                    csv_path  = os.path.join(output_dir, f"{stem}.csv")
                    xlsx_path = os.path.join(output_dir, f"{stem}.xlsx")

                    vd.write_csv(pivot_data, days, f"{approach}{movement}", csv_path)
                    self._log(f"  CSV  → {os.path.basename(csv_path)}", "ok")
                    vd.build_excel(pivot_data, days, title, xlsx_path,
                                   pivot_15min=all_pivot_15.get(movement))
                    self._log(f"  XLSX → {os.path.basename(xlsx_path)}", "ok")

            self._log("\n" + "─" * 56, "dim")
            self._log("✓  Done. All files written to output folder.", "ok")

        except Exception as e:
            self._log(f"\n✗  Error: {e}", "err")
        finally:
            self.run_btn.configure(state="normal", text="▶  Run Analysis")

    # ── Centering ────────────────────────────────────────────────────────────

    def _center(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth()  // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
