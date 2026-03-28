import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import threading
import queue
import os
import re

try:
    from PIL import Image, ImageTk as _ImageTk
    _PIL_OK = True
except ImportError:
    _PIL_OK = False

from comparison_engine import run_gui_analysis

COMMON_TERMS = ["WA", "MEDICAID"]

def compact_alias(alias: str) -> str:
    """Turn 'Q3 2025' → 'Q325', 'Q1 2026' → 'Q126', etc."""
    return re.sub(r'\s+', '', alias)   # simply drop all spaces

ACCENT   = "#2563eb"
SUCCESS  = "#16a34a"
DANGER   = "#dc2626"
SLATE    = "#475569"
BG       = "#f1f5f9"
PANEL    = "#ffffff"
TEXT     = "#1e293b"
BORDER   = "#e2e8f0"

FNORMAL  = ("Helvetica", 10)
FBOLD    = ("Helvetica", 10, "bold")

# Badge color palette — (fg/border color, light background)
BADGE_COLORS = [
    ("#2563eb", "#dbeafe"),  # blue
    ("#dc2626", "#fee2e2"),  # red
    ("#16a34a", "#dcfce7"),  # green
    ("#9333ea", "#f3e8ff"),  # purple
    ("#d97706", "#fef3c7"),  # amber
    ("#0891b2", "#cffafe"),  # cyan
    ("#db2777", "#fce7f3"),  # pink
    ("#059669", "#d1fae5"),  # emerald
]


class MatrixComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Matrix Comparator — State Rules")
        self.root.geometry("1150x740")
        self.root.minsize(900, 600)
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        # Clean exit to avoid macOS pyenv Tk segfault on close
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

        # ── App icon ──────────────────────────────────────────────────────
        _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.png")
        if os.path.exists(_icon_path):
            try:
                if _PIL_OK:
                    _img = Image.open(_icon_path).resize((256, 256), Image.LANCZOS)
                    self._icon = _ImageTk.PhotoImage(_img)
                else:
                    self._icon = tk.PhotoImage(file=_icon_path)
                self.root.iconphoto(True, self._icon)
            except Exception:
                pass   # icon is cosmetic — never crash on it

        self.src_files     = []
        self.tgt_files     = []
        self.file_sheets   = {}
        self.mapping_rows  = []
        self.log_queue     = queue.Queue()  # thread-safe message passing

        self._make_styles()
        self._build()
        self._poll_log()  # start draining log queue on main thread

    def _quit(self):
        self.root.destroy()
        os._exit(0)

    # ── TTK styles ─────────────────────────────────────────────────────────
    def _make_styles(self):
        s = ttk.Style()
        s.theme_use("clam")

        # Buttons
        s.configure("Blue.TButton",   background=ACCENT,  foreground="white",
                    font=FBOLD, padding=6, relief="flat")
        s.map("Blue.TButton",   background=[("active", "#1d4ed8"), ("disabled", "#94a3b8")])

        s.configure("Green.TButton",  background=SUCCESS, foreground="white",
                    font=FBOLD, padding=8, relief="flat")
        s.map("Green.TButton",  background=[("active", "#15803d"), ("disabled", "#94a3b8")])

        s.configure("Running.TButton", background="#b45309", foreground="white",
                    font=FBOLD, padding=8, relief="flat")
        s.map("Running.TButton", background=[("active", "#92400e"), ("disabled", "#92400e")])

        s.configure("Red.TButton",    background=DANGER,  foreground="white",
                    font=FBOLD, padding=4, relief="flat")
        s.map("Red.TButton",    background=[("active", "#b91c1c")])

        s.configure("Slate.TButton",  background=SLATE,   foreground="white",
                    font=FNORMAL, padding=5, relief="flat")
        s.map("Slate.TButton",  background=[("active", "#334155")])

        # Combobox / Entry
        s.configure("TCombobox", fieldbackground=PANEL, background=PANEL,
                    foreground=TEXT, relief="flat")
        s.configure("TEntry",    fieldbackground=PANEL, foreground=TEXT, relief="flat")
        s.configure("TScrollbar", background=BORDER, troughcolor=BG, relief="flat")

    # ── Layout helpers ─────────────────────────────────────────────────────
    def _panel(self, parent, title, **pack_kw):
        """Returns (outer_frame, inner_content_frame)."""
        outer = tk.Frame(parent, bg=PANEL, bd=1, relief="solid",
                         highlightthickness=1, highlightbackground=BORDER)
        outer.pack(**pack_kw)

        hdr = tk.Frame(outer, bg=ACCENT)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"  {title}", font=FBOLD,
                 bg=ACCENT, fg="white", anchor="w", pady=5
                 ).pack(side=tk.LEFT, fill=tk.Y)

        body = tk.Frame(outer, bg=PANEL, padx=8, pady=8)
        body.pack(fill=tk.BOTH, expand=True)
        return outer, body

    # ── Build UI ───────────────────────────────────────────────────────────
    def _build(self):
        pad = tk.Frame(self.root, bg=BG, padx=10, pady=10)
        pad.pack(fill=tk.BOTH, expand=True)

        # ── Title bar ─────────────────────────────────────────────────────
        tb = tk.Frame(pad, bg=ACCENT)
        tb.pack(fill=tk.X, pady=(0, 8))
        tk.Label(tb, text="  ⚖  Matrix Comparator — State Rules",
                 font=("Helvetica", 14, "bold"),
                 bg=ACCENT, fg="white", pady=8, anchor="w"
                 ).pack(side=tk.LEFT)

        # ── File Uploads ──────────────────────────────────────────────────
        _, files_body = self._panel(pad, "File Uploads",
                                    fill=tk.X, pady=(0, 6))

        src_col = tk.Frame(files_body, bg=PANEL)
        src_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        tk.Label(src_col, text="Source Files  (e.g. Q3, Q4)",
                 font=FBOLD, bg=PANEL, fg=TEXT).pack(anchor="w")
        # Badge wrap area for source files
        self.src_badge_frame = tk.Frame(src_col, bg=PANEL, pady=4)
        self.src_badge_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Button(src_col, text="+ Add Source Files",
                   command=self.add_src, style="Blue.TButton").pack(anchor="w", pady=(4,0))

        tk.Frame(files_body, bg=BORDER, width=1).pack(side=tk.LEFT, fill=tk.Y, padx=6)

        tgt_col = tk.Frame(files_body, bg=PANEL)
        tgt_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))
        tk.Label(tgt_col, text="Target Files  (e.g. 2026 Q1)",
                 font=FBOLD, bg=PANEL, fg=TEXT).pack(anchor="w")
        # Badge wrap area for target files
        self.tgt_badge_frame = tk.Frame(tgt_col, bg=PANEL, pady=4)
        self.tgt_badge_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Button(tgt_col, text="+ Add Target Files",
                   command=self.add_tgt, style="Blue.TButton").pack(anchor="w", pady=(4,0))


        # ── Sheet Mappings ─────────────────────────────────────────────────
        _, map_body = self._panel(pad, "Sheet Mappings",
                                   fill=tk.BOTH, expand=True, pady=(0, 6))

        # Column headers — use grid with weights so they stretch on resize
        hdr = tk.Frame(map_body, bg=PANEL)
        hdr.pack(fill=tk.X, pady=(0, 2))
        hdr.columnconfigure(0, weight=3)
        hdr.columnconfigure(1, weight=1)
        hdr.columnconfigure(2, weight=3)
        hdr.columnconfigure(3, weight=1)
        hdr.columnconfigure(4, weight=1)
        hdr.columnconfigure(5, weight=0)  # delete btn placeholder
        for col, txt in enumerate(["Source File", "Source Sheet",
                                    "Target File",  "Target Sheet",
                                    "Output Name",  ""]):
            tk.Label(hdr, text=txt, font=FBOLD, anchor="w",
                     bg=PANEL, fg=SLATE).grid(row=0, column=col,
                                              sticky="ew", padx=2)
        tk.Frame(map_body, bg=BORDER, height=1).pack(fill=tk.X, pady=(0, 4))

        # Scrollable area
        scroll_wrap = tk.Frame(map_body, bg=PANEL)
        scroll_wrap.pack(fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(scroll_wrap, orient="vertical", style="TScrollbar")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.cvas = tk.Canvas(scroll_wrap, bg=PANEL, highlightthickness=0,
                              yscrollcommand=vsb.set)
        self.cvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.configure(command=self.cvas.yview)

        self.map_cont = tk.Frame(self.cvas, bg=PANEL)
        self._map_win = self.cvas.create_window((0, 0), window=self.map_cont,
                                                anchor="nw")

        # Keep inner frame width = canvas width (horizontal fill)
        def _on_canvas_resize(e):
            self.cvas.itemconfig(self._map_win, width=e.width)
        self.cvas.bind("<Configure>", _on_canvas_resize)

        # Update scroll region whenever rows are added/removed
        self.map_cont.bind("<Configure>",
            lambda e: self.cvas.configure(scrollregion=self.cvas.bbox("all")))

        # Mousewheel / trackpad scroll (cross-platform)
        def _on_mousewheel(event):
            self.cvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _on_mousewheel_linux(event):
            self.cvas.yview_scroll(-1 if event.num == 4 else 1, "units")

        for widget in (self.cvas, self.map_cont):
            widget.bind("<MouseWheel>", _on_mousewheel)        # Windows / macOS
            widget.bind("<Button-4>",   _on_mousewheel_linux)  # Linux scroll up
            widget.bind("<Button-5>",   _on_mousewheel_linux)  # Linux scroll down

        # Add mapping button row
        add_row = tk.Frame(map_body, bg=PANEL)
        add_row.pack(fill=tk.X, pady=(6, 0))
        tk.Frame(add_row, bg=BORDER, height=1).pack(fill=tk.X, pady=(0, 4))
        ttk.Button(add_row, text="+ Add Manual Mapping",
                   command=self.add_row, style="Slate.TButton").pack(side=tk.LEFT)

        # ── Execution ──────────────────────────────────────────────────────
        _, exe_body = self._panel(pad, "Execution", fill=tk.X, pady=(0, 0))

        ctrl = tk.Frame(exe_body, bg=PANEL)
        ctrl.pack(fill=tk.X, pady=(0, 6))
        tk.Label(ctrl, text="Output File:", font=FBOLD,
                 bg=PANEL, fg=TEXT).pack(side=tk.LEFT)
        self.out_var = tk.StringVar(value="Final_Generated_Mapped_Analysis.xlsx")
        ttk.Entry(ctrl, textvariable=self.out_var, width=40).pack(
            side=tk.LEFT, padx=8)
        self.run_btn = ttk.Button(ctrl, text="▶  RUN ANALYSIS",
                                  command=self.run, style="Green.TButton")
        self.run_btn.pack(side=tk.RIGHT)

        self.log_widget = tk.Text(exe_body, height=7, state="disabled",
                           bg="#f8fafc", fg=TEXT, font=("Courier", 9),
                           relief="solid", bd=1, highlightthickness=0)
        self.log_widget.pack(fill=tk.BOTH, expand=True)

    # ── Thread-safe logger ─────────────────────────────────────────────────
    def _log(self, msg):
        """Safe to call from any thread — posts to queue, never touches widgets."""
        self.log_queue.put(msg)

    def _poll_log(self):
        """Runs only on the main thread; drains the queue and updates the Text widget."""
        try:
            while True:
                msg = self.log_queue.get_nowait()
                # Special sentinel: reset the Run button
                if msg == "__RESET_BTN__":
                    self.run_btn.config(state="normal",
                                       text="\u25b6  RUN ANALYSIS",
                                       style="Green.TButton")
                    continue
                self.log_widget.configure(state="normal")
                self.log_widget.insert(tk.END, msg + "\n")
                self.log_widget.see(tk.END)
                self.log_widget.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log)  # check again in 100 ms

    # ── File loading ───────────────────────────────────────────────────────
    def _load_sheets(self, path):
        if path in self.file_sheets:
            return
        self._log(f"Reading: {os.path.basename(path)}...")
        try:
            self.file_sheets[path] = pd.ExcelFile(path).sheet_names
            self._log(f"  → {len(self.file_sheets[path])} sheets.")
        except Exception as e:
            self._log(f"❌ {e}")

    # ── Quarter/Year picker ────────────────────────────────────────────────
    def _ask_alias(self, filepath):
        """Pop-up picker: Quarter (Q1-Q4) + Year → label like 'Q3 2025'."""
        import datetime
        current_year = datetime.date.today().year
        quarters = ["Q1", "Q2", "Q3", "Q4"]
        years    = [str(y) for y in range(2020, current_year + 5)]

        dlg = tk.Toplevel(self.root)
        dlg.title("File Period")
        dlg.resizable(False, False)
        dlg.grab_set()

        # File name label
        tk.Label(dlg, text="File:", font=FNORMAL, fg="#64748b",
                 padx=10).grid(row=0, column=0, sticky="w", padx=10, pady=(12, 0))
        tk.Label(dlg, text=os.path.basename(filepath), font=FNORMAL,
                 fg="#334155", wraplength=360, justify="left"
                 ).grid(row=0, column=1, columnspan=2, sticky="w", padx=4, pady=(12, 0))

        # Quarter picker
        tk.Label(dlg, text="Quarter:", font=FBOLD, padx=10
                 ).grid(row=1, column=0, sticky="w", padx=10, pady=(10, 4))
        qv = tk.StringVar(value="Q1")
        q_cb = ttk.Combobox(dlg, textvariable=qv, values=quarters,
                             state="readonly", width=6)
        q_cb.grid(row=1, column=1, padx=4, pady=(10, 4), sticky="w")
        q_cb.focus_set()

        # Year picker
        tk.Label(dlg, text="Year:", font=FBOLD, padx=10
                 ).grid(row=2, column=0, sticky="w", padx=10, pady=4)
        yv = tk.StringVar(value=str(current_year))
        y_cb = ttk.Combobox(dlg, textvariable=yv, values=years,
                             state="readonly", width=8)
        y_cb.grid(row=2, column=1, padx=4, pady=4, sticky="w")

        # Preview label
        preview_var = tk.StringVar()
        def _update_preview(*_):
            preview_var.set(f"Label: {qv.get()} {yv.get()}")
        qv.trace_add("write", _update_preview)
        yv.trace_add("write", _update_preview)
        _update_preview()
        tk.Label(dlg, textvariable=preview_var, font=FBOLD, fg=ACCENT,
                 padx=10).grid(row=3, column=0, columnspan=3, sticky="w",
                               padx=10, pady=(2, 8))

        result = [f"Q1 {current_year}"]
        def _ok(_=None):
            result[0] = f"{qv.get()} {yv.get()}"
            dlg.destroy()
        def _cancel(_=None):
            dlg.destroy()

        dlg.bind("<Return>", _ok)
        dlg.bind("<Escape>", _cancel)
        bf = tk.Frame(dlg)
        bf.grid(row=4, column=0, columnspan=3, pady=8)
        ttk.Button(bf, text="OK", command=_ok,
                   style="Blue.TButton", width=8).pack(side=tk.LEFT, padx=4)
        ttk.Button(bf, text="Cancel", command=_cancel,
                   style="Slate.TButton", width=8).pack(side=tk.LEFT, padx=4)

        dlg.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{x}+{y}")
        dlg.wait_window()
        return result[0]

    # ── Badge helpers ──────────────────────────────────────────────────────
    def _add_badge(self, alias, container, file_list):
        """Create a blue pill badge for `alias` inside `container`."""
        fg, bg = "#2563eb", "#dbeafe"   # always blue

        badge = tk.Frame(container, bg=bg, bd=0, pady=3, padx=3,
                         highlightbackground=fg, highlightthickness=1)
        badge.pack(side=tk.LEFT, padx=4, pady=4)

        tk.Label(badge, text=alias, bg=bg, fg=fg,
                 font=FBOLD, padx=6, pady=2).pack(side=tk.LEFT)

        def _remove():
            self._remove_file(alias, container, file_list, badge)

        close = tk.Label(badge, text=" ×", bg=bg, fg=fg,
                         font=FBOLD, cursor="hand2", padx=4, pady=2)
        close.pack(side=tk.LEFT)
        close.bind("<Button-1>", lambda _: _remove())
        close.bind("<Enter>", lambda _: close.config(fg="#ef4444"))
        close.bind("<Leave>", lambda _: close.config(fg=fg))

    def _remove_file(self, alias, container, file_list, badge_widget):
        """Remove a file by alias from file_list, clean up mappings, destroy badge."""
        # Find and remove from list
        entry = next((e for e in file_list if e[1] == alias), None)
        if entry:
            file_list.remove(entry)
            # Remove mapping rows that reference this alias
            to_kill = []
            for r in self.mapping_rows:
                if r["sfv"].get() == alias or r["tfv"].get() == alias:
                    to_kill.append(r)
            for r in to_kill:
                r["frame"].destroy()
            self.mapping_rows = [r for r in self.mapping_rows if r not in to_kill]
        badge_widget.destroy()
        self._refresh_row_combos()

    def add_src(self):
        paths = filedialog.askopenfilenames(
            title="Select Source Excel Files",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        for p in paths:
            if not any(x[0] == p for x in self.src_files):
                alias = self._ask_alias(p)
                self.src_files.append((p, alias))
                self._add_badge(alias, self.src_badge_frame, self.src_files)
                self._load_sheets(p)
        self._auto_map()
        self._refresh_row_combos()

    def add_tgt(self):
        paths = filedialog.askopenfilenames(
            title="Select Target Excel Files",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        for p in paths:
            if not any(x[0] == p for x in self.tgt_files):
                alias = self._ask_alias(p)
                self.tgt_files.append((p, alias))
                self._add_badge(alias, self.tgt_badge_frame, self.tgt_files)
                self._load_sheets(p)
        self._auto_map()
        self._refresh_row_combos()

    def _refresh_row_combos(self):
        """Update all mapping row comboboxes to reflect current aliases."""
        src_aliases = [x[1] for x in self.src_files]
        tgt_aliases = [x[1] for x in self.tgt_files]
        for r in self.mapping_rows:
            r["sf_cb"]["values"] = src_aliases
            r["tf_cb"]["values"] = tgt_aliases


    # ── Mapping rows ───────────────────────────────────────────────────────
    def add_row(self, sf="", ss="", tf="", ts="", out=""):
        row = tk.Frame(self.map_cont, bg=PANEL, pady=2)
        row.pack(fill=tk.X)

        # Proportional column widths matching the header
        row.columnconfigure(0, weight=3)
        row.columnconfigure(1, weight=1)
        row.columnconfigure(2, weight=3)
        row.columnconfigure(3, weight=1)
        row.columnconfigure(4, weight=1)
        row.columnconfigure(5, weight=0)

        sfv = tk.StringVar(value=sf)
        ssv = tk.StringVar(value=ss)
        tfv = tk.StringVar(value=tf)
        tsv = tk.StringVar(value=ts)
        ov  = tk.StringVar(value=out)

        src_names = [x[1] for x in self.src_files]
        tgt_names = [x[1] for x in self.tgt_files]

        sf_cb = ttk.Combobox(row, textvariable=sfv, values=src_names, state="readonly")
        sf_cb.grid(row=0, column=0, sticky="ew", padx=2, pady=1)
        ss_cb = ttk.Combobox(row, textvariable=ssv, state="readonly")
        ss_cb.grid(row=0, column=1, sticky="ew", padx=2, pady=1)
        tf_cb = ttk.Combobox(row, textvariable=tfv, values=tgt_names, state="readonly")
        tf_cb.grid(row=0, column=2, sticky="ew", padx=2, pady=1)
        ts_cb = ttk.Combobox(row, textvariable=tsv, state="readonly")
        ts_cb.grid(row=0, column=3, sticky="ew", padx=2, pady=1)
        ttk.Entry(row, textvariable=ov).grid(row=0, column=4, sticky="ew", padx=2, pady=1)
        del_btn = ttk.Button(row, text="✕", style="Red.TButton", width=3)
        del_btn.grid(row=0, column=5, padx=(4, 2), pady=1)

        # Propagate mousewheel to canvas so scrolling works over rows
        def _mw(e):  self.cvas.yview_scroll(int(-1*(e.delta/120)), "units")
        def _mwl(e): self.cvas.yview_scroll(-1 if e.num==4 else 1, "units")
        for w in (row, sf_cb, ss_cb, tf_cb, ts_cb, del_btn):
            w.bind("<MouseWheel>", _mw)
            w.bind("<Button-4>",   _mwl)
            w.bind("<Button-5>",   _mwl)

        def _lss(_=None):
            p = next((x[0] for x in self.src_files if x[1] == sfv.get()), None)
            if p and p in self.file_sheets:
                ss_cb["values"] = self.file_sheets[p]
                if not ss: ssv.set(self.file_sheets[p][0])
        def _lts(_=None):
            p = next((x[0] for x in self.tgt_files if x[1] == tfv.get()), None)
            if p and p in self.file_sheets:
                ts_cb["values"] = self.file_sheets[p]
                if not ts: tsv.set(self.file_sheets[p][0])

        sf_cb.bind("<<ComboboxSelected>>", _lss)
        tf_cb.bind("<<ComboboxSelected>>", _lts)
        if sf: _lss()
        if tf: _lts()

        def _del():
            row.destroy()
            self.mapping_rows = [r for r in self.mapping_rows if r["frame"] != row]
        del_btn.configure(command=_del)

        self.mapping_rows.append({
            "frame": row, "sfv": sfv, "ssv": ssv,
            "tfv": tfv, "tsv": tsv, "ov": ov,
            "sf_cb": sf_cb, "tf_cb": tf_cb
        })

    def _auto_map(self):
        if not self.src_files or not self.tgt_files:
            return
        self._log("Auto-mapping...")

        # If we now have 2+ targets, rename any existing old-format rows to the new compact format
        if len(self.tgt_files) > 1:
            src_label = "&".join(compact_alias(a) for _, a in self.src_files) if self.src_files else "Src"
            for r in self.mapping_rows:
                ov = r["ov"].get()
                tfv = r["tfv"].get()
                if " - " in ov and " vs " not in ov and tfv:
                    term_part = ov.split(" - ")[0].strip()
                    r["ov"].set(f"{term_part} - {src_label} vs {compact_alias(tfv)}")
                elif " vs " not in ov and tfv and any(ov.startswith(t) for t in COMMON_TERMS):
                    term_part = ov.split(" ")[0].strip()
                    r["ov"].set(f"{term_part} - {src_label} vs {compact_alias(tfv)}")

        # Deduplicate on (output_name, target_alias) so each target file gets its own rows
        existing = {(r["ov"].get(), r["tfv"].get()) for r in self.mapping_rows}
        count = 0

        def _match(sheet, term):
            pat = r"(?i)(^|[ \-_])" + re.escape(term) + r"($|[ \-_])"
            return bool(re.search(pat, str(sheet).strip()))

        for tp, tgt_alias in self.tgt_files:           # iterate every target file
            # Compact source label e.g. "Q325&Q425"
            src_label = "&".join(compact_alias(a) for _, a in self.src_files) if self.src_files else "Src"
            tgt_label = compact_alias(tgt_alias)
            for term in COMMON_TERMS:
                # Compact format: "MA - Q325&Q425 vs Q126"  (always ≤ 31 chars for typical aliases)
                out_n = f"{term} - {src_label} vs {tgt_label}"
                if (out_n, tgt_alias) in existing:
                    continue
                # Find matching sheet in THIS target file
                best_ts = None
                for s in self.file_sheets.get(tp, []):
                    if _match(s, term) and "updates" not in s.lower():
                        best_ts = s
                        break
                if not best_ts:
                    continue
                # Map every source file that has a matching sheet
                for sp, src_alias in self.src_files:
                    for s in self.file_sheets.get(sp, []):
                        if _match(s, term) and "updates" not in s.lower():
                            self.add_row(src_alias, s, tgt_alias, best_ts, out_n)
                            count += 1
                            break
        if count:
            self._log(f"  → {count} auto-mapping(s) created.")

    # ── Run analysis ───────────────────────────────────────────────────────
    def run(self):
        if not self.mapping_rows:
            messagebox.showwarning("Warning", "No mappings defined!")
            return

        jobs = []
        for r in self.mapping_rows:
            sfn = r["sfv"].get(); ss = r["ssv"].get()
            tfn = r["tfv"].get(); ts = r["tsv"].get()
            out = r["ov"].get()
            if not all([sfn, ss, tfn, ts]): continue
            sfp = next((p for p, n in self.src_files if n == sfn), None)
            tfp = next((p for p, n in self.tgt_files if n == tfn), None)
            if not sfp or not tfp: continue
            if not out: out = f"{ss}_vs_{ts}"[:30]
            jobs.append({"src_file": sfp, "src_sheet": ss, "src_alias": sfn,
                         "tgt_file": tfp, "tgt_sheet": ts, "tgt_alias": tfn,
                         "out_name": out})

        if not jobs:
            messagebox.showwarning("Warning", "Mappings incomplete.")
            return

        out_f = self.out_var.get()
        if not out_f.endswith(".xlsx"):
            out_f += ".xlsx"
            self.out_var.set(out_f)

        self.run_btn.config(state="disabled", text="⏳  RUNNING...", style="Running.TButton")

        def _worker():
            try:
                ok = run_gui_analysis(jobs, out_f, self._log)
                self._log("\U0001f389 Done! File saved." if ok else "\u274c Failed.")
            except Exception as e:
                self._log(f"\u274c Unexpected error: {e}")
            finally:
                # Send sentinel through the queue — _poll_log resets the button on main thread
                self.log_queue.put("__RESET_BTN__")

        threading.Thread(target=_worker, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = MatrixComparatorApp(root)
    root.mainloop()
