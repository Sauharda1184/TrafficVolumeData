"""
gui.py
Tkinter GUI for the Traffic Volume Analysis pipeline.
Run:  python3 gui.py
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Import the processing functions directly from volume_data.py
import volume_data as vd

APPROACHES = ["NB", "SB", "EB", "WB"]
APPROACH_FULL = {
    "NB": "Northbound",
    "SB": "Southbound",
    "EB": "Eastbound",
    "WB": "Westbound",
}

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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Traffic Volume Analyzer")
        self.resizable(False, False)
        self.configure(bg=BG)
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
            font=("Helvetica", 16, "bold"),
            pady=12,
        ).pack()
        tk.Label(
            hdr,
            text="Hourly Volume Analysis  •  Excel & CSV Output",
            bg=HEADER_BG, fg="#a8c8e8",
            font=("Helvetica", 9),
            pady=2,
        ).pack()

        body = tk.Frame(self, bg=BG, padx=20, pady=16)
        body.pack(fill="both")

        # Intersection name
        self._section(body, "Intersection Name")
        self.intersection_var = tk.StringVar()
        tk.Entry(
            body,
            textvariable=self.intersection_var,
            font=("Helvetica", 10),
            width=55,
            relief="solid",
            bd=1,
        ).pack(anchor="w", pady=(0, 12))

        # Approach directories
        self._section(body, "Approach Data Directories")
        tk.Label(
            body,
            text="Select the folder containing CSV files for each approach (leave blank to skip).",
            bg=BG, fg="#555", font=("Helvetica", 8, "italic"),
        ).pack(anchor="w", pady=(0, 6))

        self.dir_vars = {}
        for approach in APPROACHES:
            self._approach_row(body, approach)

        # Output directory
        self._section(body, "Output Folder")
        out_row = tk.Frame(body, bg=BG)
        out_row.pack(fill="x", pady=(0, 12))
        self.output_var = tk.StringVar(value=os.path.join(os.getcwd(), "output"))
        tk.Entry(
            out_row,
            textvariable=self.output_var,
            font=("Helvetica", 10),
            width=44,
            relief="solid",
            bd=1,
        ).pack(side="left")
        tk.Button(
            out_row,
            text="Browse",
            command=self._browse_output,
            bg=ACCENT, fg=BTN_FG,
            font=("Helvetica", 9, "bold"),
            relief="flat", padx=8, pady=2, cursor="hand2",
        ).pack(side="left", padx=(6, 0))

        # Run button
        self.run_btn = tk.Button(
            body,
            text="▶  Run Analysis",
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
            log_frame,
            height=12, width=66,
            bg=LOG_BG, fg=LOG_FG,
            font=("Courier", 9),
            state="disabled",
            relief="flat",
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
            parent,
            text=text.upper(),
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
            font=("Helvetica", 10),
            width=18, anchor="w",
        ).pack(side="left")
        var = tk.StringVar()
        self.dir_vars[approach] = var
        tk.Entry(
            row, textvariable=var,
            font=("Helvetica", 9),
            width=34, relief="solid", bd=1,
        ).pack(side="left", padx=(0, 6))
        tk.Button(
            row,
            text="Browse",
            command=lambda a=approach: self._browse_approach(a),
            bg=ACCENT, fg=BTN_FG,
            font=("Helvetica", 9, "bold"),
            relief="flat", padx=8, pady=2, cursor="hand2",
        ).pack(side="left")

    # ── Browsing ─────────────────────────────────────────────────────────────

    def _browse_approach(self, approach):
        path = filedialog.askdirectory(title=f"Select {approach} data folder")
        if path:
            self.dir_vars[approach].set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_var.set(path)

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
            messagebox.showwarning("Missing Input", "Select at least one approach directory.")
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
            self._log("─" * 54, "dim")

            for approach, directory in approach_dirs.items():
                self._log(f"\n▸ {approach} ({APPROACH_FULL.get(approach, approach)})  ← {directory}", "info")

                if not os.path.isdir(directory):
                    self._log(f"  [SKIP] Directory not found.", "err")
                    continue

                zone_map = vd.discover_zones(directory, approach)
                if not zone_map:
                    self._log(f"  [SKIP] No matching zones found.", "err")
                    continue

                for zone, mappings in zone_map.items():
                    desc = ", ".join(f"{mv}←{col}" for mv, col in mappings)
                    self._log(f"  {zone:12s}  {desc}", "dim")

                all_pivot = vd.load_all_files(directory, zone_map)

                for movement, pivot_data in all_pivot.items():
                    days  = list(pivot_data.keys())
                    label = vd.MOVEMENT_LABELS.get(movement, movement)
                    a_lbl = vd.APPROACH_LABELS.get(approach, approach)
                    title = f"{a_lbl} {label} — {intersection}"
                    stem  = f"{approach}_{movement}"

                    csv_path  = os.path.join(output_dir, f"{stem}.csv")
                    xlsx_path = os.path.join(output_dir, f"{stem}.xlsx")

                    vd.write_csv(pivot_data, days, f"{approach}{movement}", csv_path)
                    self._log(f"  CSV  → {os.path.basename(csv_path)}", "ok")

                    vd.build_excel(pivot_data, days, title, xlsx_path)
                    self._log(f"  XLSX → {os.path.basename(xlsx_path)}", "ok")

            self._log("\n" + "─" * 54, "dim")
            self._log("✓ Done. All files written to output folder.", "ok")

        except Exception as e:
            self._log(f"\n✗ Error: {e}", "err")

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
