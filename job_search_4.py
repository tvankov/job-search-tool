import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import os
import sys
import json
import subprocess
import traceback
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Config ────────────────────────────────────────────────────────────────────
_BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
_CONFIG_PATH = os.path.join(_BASE_DIR, "config.json")

def _load_config():
    if os.path.exists(_CONFIG_PATH):
        try:
            with open(_CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def _save_config(data: dict):
    existing = _load_config()
    existing.update(data)
    with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2)

_cfg    = _load_config()
APP_ID  = _cfg.get("app_id",  "")
APP_KEY = _cfg.get("app_key", "")

BG       = "#0f172a"
PANEL    = "#1e293b"
ACCENT   = "#0ea5e9"
ACCENT2  = "#38bdf8"
TEXT     = "#f1f5f9"
SUBTEXT  = "#94a3b8"
DANGER   = "#f87171"
SUCCESS  = "#34d399"
BORDER   = "#334155"
ROW_ODD  = "#1e293b"
ROW_EVEN = "#162032"

COUNTRIES = {
    "Australia":      "au",
    "Austria":        "at",
    "Belgium":        "be",
    "Brazil":         "br",
    "Canada":         "ca",
    "France":         "fr",
    "Germany":        "de",
    "India":          "in",
    "Italy":          "it",
    "Mexico":         "mx",
    "Netherlands":    "nl",
    "New Zealand":    "nz",
    "Poland":         "pl",
    "Singapore":      "sg",
    "South Africa":   "za",
    "Spain":          "es",
    "Switzerland":    "ch",
    "UK":             "gb",
    "USA":            "us",
}

SORT_OPTIONS = {
    "Relevance": "relevance",
    "Date":      "date",
    "Salary ↑":  "salary_asc",
    "Salary ↓":  "salary_desc",
}

RESULTS_OPTIONS = [5, 10, 20, 50]

# ── Main App ──────────────────────────────────────────────────────────────────
class JobSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Job Search Tool  —  by Todor Vankov")
        self.geometry("980x680")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.jobs = []
        self._style()
        self._header()
        self._footer()
        self._notebook()

    # ── Style ─────────────────────────────────────────────────────────────────
    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",       background=BG)
        s.configure("TLabel",       background=BG, foreground=TEXT, font=("Segoe UI", 10))
        s.configure("TCombobox",    fieldbackground=PANEL, background=PANEL,
                    foreground=TEXT, selectbackground=ACCENT, font=("Segoe UI", 10))
        s.map("TCombobox",
              fieldbackground=[("readonly", PANEL), ("disabled", PANEL)],
              selectbackground=[("readonly", PANEL)],
              selectforeground=[("readonly", TEXT)])
        s.configure("Treeview",     background=PANEL, foreground=TEXT,
                    fieldbackground=PANEL, rowheight=28, font=("Segoe UI", 10))
        s.configure("Treeview.Heading", background=ACCENT, foreground="white",
                    font=("Segoe UI", 10, "bold"), relief="flat")
        s.map("Treeview", background=[("selected", ACCENT)])
        s.configure("TScrollbar",   background=PANEL, troughcolor=BG,
                    bordercolor=BORDER, arrowcolor=ACCENT2)
        s.configure("TNotebook",        background=BG, borderwidth=0)
        s.configure("TNotebook.Tab",    background=PANEL, foreground=SUBTEXT,
                    font=("Segoe UI", 10), padding=(14, 6))
        s.map("TNotebook.Tab",
              background=[("selected", BG)],
              foreground=[("selected", ACCENT2)])

    # ── Header ────────────────────────────────────────────────────────────────
    def _header(self):
        bar = tk.Frame(self, bg=PANEL, height=54)
        bar.pack(fill="x")
        tk.Label(bar, text="🔍  Job Search Tool", bg=PANEL, fg=ACCENT2,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=20, pady=12)

    # ── Notebook ──────────────────────────────────────────────────────────────
    def _notebook(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=0, pady=0)

        self.tab_search      = tk.Frame(self.nb, bg=BG)
        self.tab_saved       = tk.Frame(self.nb, bg=BG)
        self.tab_autorun     = tk.Frame(self.nb, bg=BG)
        self.tab_credentials = tk.Frame(self.nb, bg=BG)
        self.tab_help        = tk.Frame(self.nb, bg=BG)
        self.tab_about       = tk.Frame(self.nb, bg=BG)
        self.nb.add(self.tab_search,      text="  Search  ")
        self.nb.add(self.tab_saved,       text="  Saved Results  ")
        self.nb.add(self.tab_autorun,     text="  Auto Run  ")
        self.nb.add(self.tab_credentials, text="  Credentials  ")
        self.nb.add(self.tab_help,        text="  Help  ")
        self.nb.add(self.tab_about,       text="  About  ")

        self._search_panel()
        self._results_panel()
        self._saved_panel()
        self._settings_panel()
        self._credentials_panel()
        self._help_panel()
        self._about_panel()
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

    # ── Search Panel ──────────────────────────────────────────────────────────
    def _search_panel(self):
        panel = tk.Frame(self.tab_search, bg=PANEL, pady=14)
        panel.pack(fill="x")

        row1 = tk.Frame(panel, bg=PANEL)
        row1.pack(fill="x", padx=20, pady=(0, 8))

        tk.Label(row1, text="Job Title", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.what_var = tk.StringVar(value="Data Analyst")
        tk.Entry(row1, textvariable=self.what_var, bg=BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 11),
                 width=22).grid(row=1, column=0, padx=(0, 12), ipady=6)

        tk.Label(row1, text="Location", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=1, sticky="w")
        self.where_var = tk.StringVar(value="Berlin")
        tk.Entry(row1, textvariable=self.where_var, bg=BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 11),
                 width=18).grid(row=1, column=1, padx=(0, 12), ipady=6)

        tk.Label(row1, text="Country", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=2, sticky="w")
        self.country_var = tk.StringVar(value="Germany")
        ttk.Combobox(row1, textvariable=self.country_var, values=list(COUNTRIES.keys()),
                     width=16, state="readonly").grid(row=1, column=2, padx=(0, 12), ipady=4)

        tk.Label(row1, text="Results", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=3, sticky="w")
        self.results_var = tk.IntVar(value=10)
        ttk.Combobox(row1, textvariable=self.results_var, values=RESULTS_OPTIONS,
                     width=6, state="readonly").grid(row=1, column=3, padx=(0, 12), ipady=4)

        tk.Label(row1, text="Sort by", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=4, sticky="w")
        self.sort_var = tk.StringVar(value="Relevance")
        ttk.Combobox(row1, textvariable=self.sort_var, values=list(SORT_OPTIONS.keys()),
                     width=12, state="readonly").grid(row=1, column=4, padx=(0, 12), ipady=4)

        self._btn(row1, "🔍  Search", self._search, w=14).grid(row=1, column=5, padx=(8, 0))

        row2 = tk.Frame(panel, bg=PANEL)
        row2.pack(fill="x", padx=20)

        tk.Label(row2, text="Min. Salary (€)", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.salary_min_var = tk.StringVar()
        tk.Entry(row2, textvariable=self.salary_min_var, bg=BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 10),
                 width=12).grid(row=1, column=0, padx=(0, 12), ipady=4)

        tk.Label(row2, text="Max. Salary (€)", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=1, sticky="w")
        self.salary_max_var = tk.StringVar()
        tk.Entry(row2, textvariable=self.salary_max_var, bg=BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 10),
                 width=12).grid(row=1, column=1, padx=(0, 12), ipady=4)

        tk.Label(row2, text="Full-time only", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=2, sticky="w")
        self.fulltime_var = tk.BooleanVar()
        tk.Checkbutton(row2, variable=self.fulltime_var, bg=PANEL, fg=TEXT,
                       selectcolor=BG, activebackground=PANEL).grid(row=1, column=2, padx=(0, 12), sticky="w")

        tk.Label(row2, text="Permanent only", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=3, sticky="w")
        self.permanent_var = tk.BooleanVar()
        tk.Checkbutton(row2, variable=self.permanent_var, bg=PANEL, fg=TEXT,
                       selectcolor=BG, activebackground=PANEL).grid(row=1, column=3, padx=(0, 12), sticky="w")

        tk.Label(row2, text="Remote only", bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).grid(row=0, column=4, sticky="w")
        self.remote_var = tk.BooleanVar()
        tk.Checkbutton(row2, variable=self.remote_var, bg=PANEL, fg=TEXT,
                       selectcolor=BG, activebackground=PANEL).grid(row=1, column=4, sticky="w")

    # ── Results Panel ─────────────────────────────────────────────────────────
    def _results_panel(self):
        frame = tk.Frame(self.tab_search, bg=BG)
        frame.pack(fill="both", expand=True, padx=16, pady=10)

        cols = ("title", "company", "location", "salary", "date", "url")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")

        for col, (label, width) in {
            "title":    ("Job Title", 320),
            "company":  ("Company",   160),
            "location": ("Location",  130),
            "salary":   ("Salary",     90),
            "date":     ("Posted",     90),
            "url":      ("Link",      120),
        }.items():
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, anchor="w")

        self.tree.tag_configure("odd",  background=ROW_ODD)
        self.tree.tag_configure("even", background=ROW_EVEN)

        sb_y = ttk.Scrollbar(frame, orient="vertical",   command=self.tree.yview)
        sb_x = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._open_link)

    # ── Saved Results Panel ───────────────────────────────────────────────────
    def _saved_panel(self):
        pane = tk.PanedWindow(self.tab_saved, orient="horizontal",
                              bg=BG, sashwidth=6, sashrelief="flat")
        pane.pack(fill="both", expand=True)

        # ── Left: file list ──────────────────────────────────────────────────
        left = tk.Frame(pane, bg=PANEL, width=220)
        pane.add(left, minsize=160)

        tk.Label(left, text="Saved Files", bg=PANEL, fg=ACCENT2,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(12, 6))

        list_frame = tk.Frame(left, bg=PANEL)
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        sb = tk.Scrollbar(list_frame, bg=PANEL, troughcolor=BG, relief="flat")
        sb.pack(side="right", fill="y")

        self.file_listbox = tk.Listbox(
            list_frame, bg=BG, fg=TEXT, selectbackground=ACCENT,
            selectforeground="white", font=("Segoe UI", 10),
            relief="flat", borderwidth=0, activestyle="none",
            yscrollcommand=sb.set)
        self.file_listbox.pack(fill="both", expand=True)
        sb.config(command=self.file_listbox.yview)
        self.file_listbox.bind("<<ListboxSelect>>", self._load_saved_file)

        btn_left = tk.Frame(left, bg=PANEL)
        btn_left.pack(pady=(0, 10))
        self._btn(btn_left, "Refresh", self._refresh_saved_list,
                  color="#334155", w=10).pack(side="left", padx=(8, 4))
        self._btn(btn_left, "Delete", self._delete_saved_file,
                  color="#7f1d1d", w=8).pack(side="left", padx=(0, 8))

        # ── Right: job table ─────────────────────────────────────────────────
        right = tk.Frame(pane, bg=BG)
        pane.add(right, minsize=400)

        self.saved_info = tk.Label(right, text="Select a file to view its contents.",
                                   bg=BG, fg=SUBTEXT, font=("Segoe UI", 9))
        self.saved_info.pack(anchor="w", padx=12, pady=(10, 4))

        cols = ("title", "company", "location", "salary_min", "salary_max", "posted", "url")
        self.saved_tree = ttk.Treeview(right, columns=cols, show="headings", selectmode="browse")

        for col, (label, width) in {
            "title":      ("Job Title",   300),
            "company":    ("Company",     150),
            "location":   ("Location",    120),
            "salary_min": ("Salary Min",   90),
            "salary_max": ("Salary Max",   90),
            "posted":     ("Posted",       90),
            "url":        ("Link",        120),
        }.items():
            self.saved_tree.heading(col, text=label)
            self.saved_tree.column(col, width=width, anchor="w")

        self.saved_tree.tag_configure("odd",  background=ROW_ODD)
        self.saved_tree.tag_configure("even", background=ROW_EVEN)

        sb_y = ttk.Scrollbar(right, orient="vertical",   command=self.saved_tree.yview)
        sb_x = ttk.Scrollbar(right, orient="horizontal", command=self.saved_tree.xview)
        self.saved_tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.saved_tree.pack(fill="both", expand=True, padx=(8, 0), pady=(0, 8))
        self.saved_tree.bind("<Double-1>", self._open_saved_link)

        self._saved_files = {}  # display_name -> full_path
        self._refresh_saved_list()

    def _refresh_saved_list(self):
        self._saved_files.clear()
        self.file_listbox.delete(0, "end")
        jobs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jobs")
        if not os.path.isdir(jobs_dir):
            return
        for root, _, files in os.walk(jobs_dir):
            for f in files:
                if f.endswith(".xlsx"):
                    full = os.path.join(root, f)
                    label = os.path.splitext(f)[0].replace("_", " ")
                    self._saved_files[label] = full
        for label in sorted(self._saved_files):
            self.file_listbox.insert("end", label)

    def _load_saved_file(self, event=None):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        label = self.file_listbox.get(sel[0])
        path  = self._saved_files.get(label)
        if not path or not os.path.exists(path):
            return

        for row in self.saved_tree.get_children():
            self.saved_tree.delete(row)

        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            if "Job Results" not in wb.sheetnames:
                self.saved_info.config(text="No 'Job Results' sheet found in this file.")
                return
            ws = wb["Job Results"]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            for i, r in enumerate(rows):
                title, company, location, sal_min, sal_max, posted, url = (
                    (r[0] or ""), (r[1] or ""), (r[2] or ""),
                    (r[3] or ""), (r[4] or ""), (r[5] or ""), (r[6] or ""))
                self.saved_tree.insert("", "end", iid=str(i),
                                       tags=("odd" if i % 2 else "even",),
                                       values=(title, company, location,
                                               sal_min, sal_max, posted, url))
            wb.close()
            self.saved_info.config(
                text=f"{label}  —  {len(rows)} jobs  |  {path}", fg=SUBTEXT)
        except Exception as e:
            self.saved_info.config(text=f"Error reading file: {e}", fg=DANGER)

    def _delete_saved_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        label = self.file_listbox.get(sel[0])
        path  = self._saved_files.get(label)
        if not path or not os.path.exists(path):
            return
        if not messagebox.askyesno("Delete File",
                                   f"Delete '{label}'?\n\n{path}\n\nThis cannot be undone."):
            return
        try:
            os.remove(path)
            folder = os.path.dirname(path)
            if not os.listdir(folder):
                os.rmdir(folder)
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete file:\n{e}")
            return
        for row in self.saved_tree.get_children():
            self.saved_tree.delete(row)
        self.saved_info.config(text="Select a file to view its contents.", fg=SUBTEXT)
        self._refresh_saved_list()

    def _open_saved_link(self, event=None):
        sel = self.saved_tree.selection()
        if sel:
            url = self.saved_tree.item(sel[0])["values"][6]
            if url:
                __import__("webbrowser").open(str(url))

    # ── Auto Run Panel ────────────────────────────────────────────────────────
    def _settings_panel(self):
        outer = tk.Frame(self.tab_autorun, bg=BG)
        outer.pack(fill="both", expand=True, padx=30, pady=24)

        # Schedule section
        section = tk.Frame(outer, bg=PANEL, pady=18)
        section.pack(fill="x")

        tk.Label(section, text="Scheduled Auto-Run", bg=PANEL, fg=ACCENT2,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(0, 4))
        tk.Label(section, text="Run the job scraper automatically every day at a set time (Windows Task Scheduler).",
                 bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).pack(anchor="w", padx=20, pady=(0, 14))

        time_row = tk.Frame(section, bg=PANEL)
        time_row.pack(anchor="w", padx=20, pady=(0, 14))

        tk.Label(time_row, text="Time (HH:MM)", bg=PANEL, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.sched_time_var = tk.StringVar(value="08:00")
        tk.Entry(time_row, textvariable=self.sched_time_var, bg=BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 11),
                 width=10).grid(row=1, column=0, ipady=6, padx=(0, 16))

        btn_row = tk.Frame(section, bg=PANEL)
        btn_row.pack(anchor="w", padx=20)
        self._btn(btn_row, "Schedule Daily", self._schedule, w=16).pack(side="left", padx=(0, 10))
        self._btn(btn_row, "Remove Schedule", self._unschedule, color="#7f1d1d", w=16).pack(side="left")

        # ── Search History ───────────────────────────────────────────────────
        tk.Frame(outer, bg=BORDER, height=1).pack(fill="x", pady=(20, 0))

        hist_header = tk.Frame(outer, bg=BG)
        hist_header.pack(fill="x", pady=(12, 6))
        tk.Label(hist_header, text="Search History  (last 30)", bg=BG, fg=ACCENT2,
                 font=("Segoe UI", 11, "bold")).pack(side="left")
        self._btn(hist_header, "Refresh", self._refresh_history,
                  color="#334155", w=10).pack(side="right")

        hist_cols = ("datetime", "query", "location", "fetched", "new", "skipped")
        self.hist_tree = ttk.Treeview(outer, columns=hist_cols, show="headings",
                                      selectmode="none", height=12)
        for col, (label, width) in {
            "datetime": ("Date / Time",  140),
            "query":    ("Search",       160),
            "location": ("Location",     110),
            "fetched":  ("Fetched",       70),
            "new":      ("New",           60),
            "skipped":  ("Skipped",       70),
        }.items():
            self.hist_tree.heading(col, text=label)
            self.hist_tree.column(col, width=width, anchor="w")

        self.hist_tree.tag_configure("odd",  background=ROW_ODD)
        self.hist_tree.tag_configure("even", background=ROW_EVEN)

        hist_sb = ttk.Scrollbar(outer, orient="vertical", command=self.hist_tree.yview)
        self.hist_tree.configure(yscrollcommand=hist_sb.set)
        hist_sb.pack(side="right", fill="y")
        self.hist_tree.pack(fill="both", expand=True, pady=(0, 10))

        self._refresh_history()

    def _parse_log_history(self):
        """Return list of dicts for each individual search entry in the log (newest first)."""
        import re
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "job_search.log")
        if not os.path.exists(log_path):
            return []

        entries = []
        current_dt = None
        pattern_search  = re.compile(r"Searching: '(.+)' in '(.+)' \(")
        pattern_fetched = re.compile(r">> (\d+) jobs fetched")
        pattern_result  = re.compile(r">> (\d+) new\s+\|\s+(\d+) duplicates skipped")
        pattern_time    = re.compile(r"\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]")

        pending = None
        with open(log_path, encoding="utf-8") as f:
            for line in f:
                m_time = pattern_time.match(line)
                if m_time:
                    current_dt = m_time.group(1)
                m_search = pattern_search.search(line)
                if m_search:
                    if pending:
                        entries.append(pending)
                    pending = {"datetime": current_dt,
                               "query": m_search.group(1),
                               "location": m_search.group(2),
                               "fetched": "", "new": "", "skipped": ""}
                    continue
                if pending:
                    m_f = pattern_fetched.search(line)
                    if m_f:
                        pending["fetched"] = m_f.group(1)
                        continue
                    m_r = pattern_result.search(line)
                    if m_r:
                        pending["new"]     = m_r.group(1)
                        pending["skipped"] = m_r.group(2)
                        entries.append(pending)
                        pending = None
        if pending:
            entries.append(pending)
        return list(reversed(entries))

    def _refresh_history(self):
        for row in self.hist_tree.get_children():
            self.hist_tree.delete(row)
        entries = self._parse_log_history()[:30]
        for i, e in enumerate(entries):
            self.hist_tree.insert("", "end", tags=("odd" if i % 2 else "even",),
                                  values=(e["datetime"], e["query"], e["location"],
                                          e["fetched"], e["new"], e["skipped"]))

    def _on_tab_change(self, _event=None):
        tab = self.nb.select()
        if tab == str(self.tab_saved):
            self._refresh_saved_list()
        elif tab == str(self.tab_autorun):
            self._refresh_history()

    # ── Credentials Panel ─────────────────────────────────────────────────────
    def _credentials_panel(self):
        outer = tk.Frame(self.tab_credentials, bg=BG)
        outer.pack(fill="both", expand=True, padx=30, pady=24)

        section = tk.Frame(outer, bg=PANEL, pady=18)
        section.pack(fill="x")

        tk.Label(section, text="API Credentials", bg=PANEL, fg=ACCENT2,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(0, 4))
        tk.Label(section,
                 text="Enter your Adzuna API credentials below. Get them for free at adzuna.com.",
                 bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9)).pack(anchor="w", padx=20, pady=(0, 18))

        fields = tk.Frame(section, bg=PANEL)
        fields.pack(anchor="w", padx=20, pady=(0, 18))

        tk.Label(fields, text="App ID", bg=PANEL, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.app_id_var = tk.StringVar(value=APP_ID)
        tk.Entry(fields, textvariable=self.app_id_var, bg=BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 11),
                 width=28).grid(row=1, column=0, ipady=6, padx=(0, 20))

        tk.Label(fields, text="App Key", bg=PANEL, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(row=0, column=1, sticky="w")
        self.app_key_var = tk.StringVar(value=APP_KEY)
        tk.Entry(fields, textvariable=self.app_key_var, bg=BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 11),
                 width=36).grid(row=1, column=1, ipady=6)

        btn_row = tk.Frame(section, bg=PANEL)
        btn_row.pack(anchor="w", padx=20)
        self._btn(btn_row, "Save Credentials", self._save_credentials, w=18).pack(side="left")

        self.cred_status = tk.Label(section, text="", bg=PANEL, fg=SUCCESS,
                                    font=("Segoe UI", 9))
        self.cred_status.pack(anchor="w", padx=20, pady=(10, 0))

        # ── Tutorial ─────────────────────────────────────────────────────────
        tk.Frame(outer, bg=BORDER, height=1).pack(fill="x", pady=(20, 0))

        tut = tk.Frame(outer, bg=PANEL, pady=18)
        tut.pack(fill="x", pady=(12, 0))

        tk.Label(tut, text="How to get your free API credentials", bg=PANEL, fg=ACCENT2,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=20, pady=(0, 12))

        steps = [
            ("1", "Go to developer.adzuna.com and click  \"Register\""),
            ("2", "Create a free account with your e-mail address"),
            ("3", "After login, click  \"Create new application\""),
            ("4", "Fill in any app name (e.g. \"Job Search\") and click  \"Save\""),
            ("5", "Copy the  App ID  and  App Key  shown on the dashboard"),
            ("6", "Paste them into the fields above and click  \"Save Credentials\""),
        ]

        for num, text in steps:
            row = tk.Frame(tut, bg=PANEL)
            row.pack(anchor="w", padx=20, pady=2)
            tk.Label(row, text=num, bg=ACCENT, fg="white",
                     font=("Segoe UI", 8, "bold"), width=2,
                     relief="flat").pack(side="left", padx=(0, 10))
            tk.Label(row, text=text, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 10)).pack(side="left")

        link_row = tk.Frame(tut, bg=PANEL)
        link_row.pack(anchor="w", padx=20, pady=(14, 4))
        tk.Label(link_row, text="Link: ", bg=PANEL, fg=SUBTEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        lnk = tk.Label(link_row, text="developer.adzuna.com", bg=PANEL, fg=ACCENT2,
                       font=("Segoe UI", 9, "underline"), cursor="hand2")
        lnk.pack(side="left")
        lnk.bind("<Button-1>", lambda e: __import__("webbrowser").open("https://developer.adzuna.com"))

    def _save_credentials(self):
        global APP_ID, APP_KEY
        new_id  = self.app_id_var.get().strip()
        new_key = self.app_key_var.get().strip()
        if not new_id or not new_key:
            self.cred_status.config(text="App ID and App Key cannot be empty.", fg=DANGER)
            return
        APP_ID  = new_id
        APP_KEY = new_key
        try:
            _save_config({"app_id": new_id, "app_key": new_key})
            self.cred_status.config(text="Credentials saved and will be remembered next time.", fg=SUCCESS)
        except Exception as e:
            self.cred_status.config(text=f"Saved for this session (file error: {e})", fg=ACCENT2)

    # ── Help Panel ────────────────────────────────────────────────────────────
    def _help_panel(self):
        # Scrollable canvas
        canvas = tk.Canvas(self.tab_help, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(self.tab_help, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=BG)
        win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_resize(e):
            canvas.itemconfig(win, width=e.width)
        def _on_frame(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.bind("<Configure>", _on_resize)
        inner.bind("<Configure>", _on_frame)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        P = 30  # horizontal padding

        def section(title):
            tk.Frame(inner, bg=BORDER, height=1).pack(fill="x", padx=P, pady=(24, 0))
            tk.Label(inner, text=title, bg=BG, fg=ACCENT2,
                     font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=P, pady=(10, 6))

        def step(num, text):
            row = tk.Frame(inner, bg=BG)
            row.pack(anchor="w", padx=P, pady=3)
            tk.Label(row, text=str(num), bg=ACCENT, fg="white",
                     font=("Segoe UI", 8, "bold"), width=2,
                     relief="flat").pack(side="left", padx=(0, 12))
            tk.Label(row, text=text, bg=BG, fg=TEXT,
                     font=("Segoe UI", 10), justify="left").pack(side="left")

        def info(label, text):
            row = tk.Frame(inner, bg=BG)
            row.pack(anchor="w", padx=P, pady=3)
            tk.Label(row, text=label, bg=BG, fg=ACCENT2,
                     font=("Segoe UI", 10, "bold"), width=18,
                     anchor="w").pack(side="left")
            tk.Label(row, text=text, bg=BG, fg=TEXT,
                     font=("Segoe UI", 10), justify="left").pack(side="left")

        def faq(question, answer):
            tk.Label(inner, text=question, bg=BG, fg=TEXT,
                     font=("Segoe UI", 10, "bold"), anchor="w").pack(
                         anchor="w", padx=P, pady=(10, 1))
            tk.Label(inner, text=answer, bg=BG, fg=SUBTEXT,
                     font=("Segoe UI", 10), anchor="w", justify="left").pack(
                         anchor="w", padx=P + 12)

        # ── Title ─────────────────────────────────────────────────────────────
        tk.Label(inner, text="How to use Job Search Tool", bg=BG, fg=TEXT,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=P, pady=(24, 2))
        tk.Label(inner, text="Find and save job listings from Adzuna — automatically, every day.",
                 bg=BG, fg=SUBTEXT, font=("Segoe UI", 10)).pack(anchor="w", padx=P)

        # ── Quick Start ───────────────────────────────────────────────────────
        section("Quick Start  —  your first search in 3 steps")
        step(1, 'Go to the Credentials tab and enter your Adzuna App ID and App Key, then click "Save Credentials".')
        step(2, 'Open the Search tab, type a Job Title (e.g. "Data Analyst") and a Location (e.g. "Berlin").')
        step(3, 'Click Search. Results appear instantly. Double-click any row to open the job listing.')

        # ── Tab Guide ─────────────────────────────────────────────────────────
        section("Tab Guide")
        info("Search",         "Search for jobs live. Results are shown in the table below the filters.\n"
                               "                   A new Excel file is created automatically on the first search.")
        info("Saved Results",  "Browse all your saved Excel files. Click a file to view its jobs.\n"
                               "                   Double-click a row to open the listing. Use Delete to remove a file.")
        info("Auto Run",       "Schedule the scraper to run every day at a set time (Windows only).\n"
                               "                   It uses the Job Title and Location you searched last.")
        info("Credentials",    "Enter and save your Adzuna API keys. They are stored in config.json\n"
                               "                   and loaded automatically on every start.")
        info("Help",           "This page.")

        # ── Auto Run guide ────────────────────────────────────────────────────
        section("Setting up Auto Run")
        step(1, "Run a search with the Job Title and Location you want to track daily.")
        step(2, 'Go to the Auto Run tab and set the time (e.g. "08:00").')
        step(3, 'Click "Schedule Daily". Windows Task Scheduler will run the scraper every day at that time.')
        step(4, "New jobs are appended to your Excel file. Duplicates are skipped automatically.")
        step(5, 'To stop: click "Remove Schedule".')

        # ── FAQ ───────────────────────────────────────────────────────────────
        section("FAQ")
        faq("Where are my Excel files saved?",
            "In a  jobs/  folder next to the app, organised by job title.\n"
            "Example:  jobs/Data_Analyst/Data_Analyst.xlsx")
        faq("Why do I get a 401 error?",
            "Your App ID or App Key is wrong or missing.\n"
            "Go to the Credentials tab, re-enter your keys and click Save Credentials.")
        faq("Can I search multiple locations?",
            "Yes — run a search for each location. Each result is appended\n"
            "to the same Excel file for that job title.")
        faq("The scheduled run found 0 jobs — what happened?",
            "Check job_search.log in the app folder for details.\n"
            "Most common cause: invalid credentials or no internet connection.")
        faq("How do I update the scheduled search?",
            "Run a new search with the updated Job Title / Location,\n"
            'then click "Schedule Daily" again — it overwrites the previous setting.')

        tk.Frame(inner, bg=BG, height=30).pack()  # bottom padding

    # ── About Panel ───────────────────────────────────────────────────────────
    def _about_panel(self):
        outer = tk.Frame(self.tab_about, bg=BG)
        outer.pack(fill="both", expand=True)

        center = tk.Frame(outer, bg=BG)
        center.place(relx=0.5, rely=0.5, anchor="center")

        # Logo / App name
        tk.Label(center, text="Job Search Tool", bg=BG, fg=TEXT,
                 font=("Segoe UI", 26, "bold")).pack(pady=(0, 4))
        tk.Label(center, text="Version 1.0", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 10)).pack(pady=(0, 24))

        # Divider
        tk.Frame(center, bg=BORDER, height=1, width=380).pack(pady=(0, 24))

        # Author
        tk.Label(center, text="Built by", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 10)).pack()
        name_lbl = tk.Label(center, text="Todor Vankov", bg=BG, fg=ACCENT2,
                            font=("Segoe UI", 14, "bold underline"), cursor="hand2")
        name_lbl.pack(pady=(2, 2))
        name_lbl.bind("<Button-1>",
                      lambda e: __import__("webbrowser").open("https://www.todorvankov.com"))

        site_lbl = tk.Label(center, text="www.todorvankov.com", bg=BG, fg=SUBTEXT,
                            font=("Segoe UI", 9, "underline"), cursor="hand2")
        site_lbl.pack(pady=(0, 24))
        site_lbl.bind("<Button-1>",
                      lambda e: __import__("webbrowser").open("https://www.todorvankov.com"))

        # Divider
        tk.Frame(center, bg=BORDER, height=1, width=380).pack(pady=(0, 24))

        # Tech stack
        tk.Label(center, text="Built with", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).pack(pady=(0, 8))

        stack_row = tk.Frame(center, bg=BG)
        stack_row.pack(pady=(0, 24))
        for tech, color in [("Python", "#3b82f6"), ("Tkinter", "#8b5cf6"),
                             ("openpyxl", "#10b981"), ("Adzuna API", "#f59e0b")]:
            tk.Label(stack_row, text=tech, bg=color, fg="white",
                     font=("Segoe UI", 9, "bold"),
                     padx=10, pady=4, relief="flat").pack(side="left", padx=4)

        # Divider
        tk.Frame(center, bg=BORDER, height=1, width=380).pack(pady=(0, 16))

        # Data source
        tk.Label(center, text="Job data provided by Adzuna  —  adzuna.com",
                 bg=BG, fg=SUBTEXT, font=("Segoe UI", 9)).pack()
        tk.Label(center, text="Free API — up to 250 requests / month on the basic plan",
                 bg=BG, fg=SUBTEXT, font=("Segoe UI", 9)).pack(pady=(2, 0))

    # ── Footer ────────────────────────────────────────────────────────────────
    def _footer(self):
        bar = tk.Frame(self, bg=PANEL, pady=8)
        bar.pack(fill="x", side="bottom")

        self.status_lbl = tk.Label(bar, text="Ready — enter a search and click 🔍 Search",
                                   bg=PANEL, fg=SUBTEXT, font=("Segoe UI", 9))
        self.status_lbl.pack(side="left", padx=16)

        self._btn(bar, "💾  Export to Excel", self._export_excel, w=20).pack(side="right", padx=12)
        self._btn(bar, "🗑  Clear Results",   self._clear,        color="#334155", w=14).pack(side="right", padx=(0, 4))

    # ── Helper ────────────────────────────────────────────────────────────────
    def _btn(self, parent, text, cmd, color=ACCENT, w=None):
        kw = dict(text=text, command=cmd, bg=color, fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat",
                  cursor="hand2", activebackground=ACCENT2,
                  activeforeground="white", pady=6, padx=12)
        if w:
            kw["width"] = w
        b = tk.Button(parent, **kw)
        b.bind("<Enter>", lambda e: b.config(bg=ACCENT2))
        b.bind("<Leave>", lambda e: b.config(bg=color))
        return b

    def _set_status(self, msg, ok=True):
        self.status_lbl.config(text=msg, fg=SUCCESS if ok else DANGER)

    # ── Search ────────────────────────────────────────────────────────────────
    def _search(self):
        country = COUNTRIES.get(self.country_var.get(), "de")
        sort_by = SORT_OPTIONS.get(self.sort_var.get(), "relevance")

        params = {
            "app_id":           self.app_id_var.get().strip(),
            "app_key":          self.app_key_var.get().strip(),
            "what":             self.what_var.get().strip(),
            "where":            self.where_var.get().strip(),
            "results_per_page": self.results_var.get(),
            "sort_by":          sort_by,
        }
        if self.salary_min_var.get().strip(): params["salary_min"] = self.salary_min_var.get().strip()
        if self.salary_max_var.get().strip(): params["salary_max"] = self.salary_max_var.get().strip()
        if self.fulltime_var.get():           params["full_time"]  = 1
        if self.permanent_var.get():          params["permanent"]  = 1

        self._set_status("Searching...")
        self.update()

        try:
            resp = requests.get(
                f"https://api.adzuna.com/v1/api/jobs/{country}/search/1",
                params=params, timeout=10)

            if resp.status_code != 200:
                self._set_status(f"API Error {resp.status_code}", ok=False)
                return

            data       = resp.json()
            results    = data.get("results", [])
            self.jobs  = results

            for row in self.tree.get_children():
                self.tree.delete(row)

            for i, job in enumerate(results):
                sal_min = job.get("salary_min")
                sal_max = job.get("salary_max")
                salary  = (f"{int(sal_min):,} – {int(sal_max):,}" if sal_min and sal_max
                           else f"from {int(sal_min):,}" if sal_min else "")
                date_raw = job.get("created", "")
                self.tree.insert("", "end", iid=str(i),
                                 tags=("odd" if i % 2 else "even",), values=(
                    job.get("title", ""),
                    job.get("company", {}).get("display_name", ""),
                    job.get("location", {}).get("display_name", ""),
                    salary, date_raw[:10],
                    job.get("redirect_url", ""),
                ))

            total = data.get("count", len(results))
            self._set_status(f"✓ {len(results)} jobs found  (total on Adzuna: {total:,})")

            job_title = self.what_var.get().strip().replace(" ", "_")
            excel_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "jobs", job_title,
                f"{job_title}.xlsx")
            if not os.path.exists(excel_path):
                self._export_excel(silent=True)

        except requests.exceptions.ConnectionError:
            self._set_status("⚠ No internet connection", ok=False)
        except Exception as e:
            self._set_status(f"Error: {e}", ok=False)

    # ── Double click → open URL ───────────────────────────────────────────────
    def _open_link(self, event):
        sel = self.tree.selection()
        if sel:
            url = self.tree.item(sel[0])["values"][5]
            if url:
                __import__("webbrowser").open(url)

    # ── Clear ─────────────────────────────────────────────────────────────────
    def _clear(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.jobs = []
        self._set_status("Results cleared")

    # ── Export to Excel ───────────────────────────────────────────────────────
    def _export_excel(self, silent=False):
        if not self.jobs:
            if not silent:
                messagebox.showwarning("No data", "Please run a search first.")
            return

        job_title = self.what_var.get().strip().replace(" ", "_")
        base_dir  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jobs", job_title)
        os.makedirs(base_dir, exist_ok=True)
        path = os.path.join(base_dir, f"{job_title}.xlsx")

        headers     = ["Title", "Company", "Location", "Salary Min (€)", "Salary Max (€)",
                       "Posted", "Link", "Description", "Imported On"]
        hdr_fill    = PatternFill("solid", start_color="0ea5e9")
        hdr_font    = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        hdr_align   = Alignment(horizontal="center", vertical="center")
        thin_border = Border(bottom=Side(style="thin", color="334155"))
        odd_fill    = PatternFill("solid", start_color="1e293b")
        even_fill   = PatternFill("solid", start_color="162032")
        data_font   = Font(name="Arial", size=10, color="F1F5F9")
        link_font   = Font(name="Arial", size=10, color="38bdf8", underline="single")

        if os.path.exists(path):
            wb = load_workbook(path)
            ws = wb["Job Results"]
            existing_urls = {str(r[6]) for r in ws.iter_rows(min_row=2, values_only=True) if r[6]}
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Job Results"
            existing_urls = set()
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = hdr_font; cell.fill = hdr_fill
                cell.alignment = hdr_align; cell.border = thin_border
            ws.row_dimensions[1].height = 22

        next_row = ws.max_row + 1
        added = skipped = 0
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        for job in self.jobs:
            url = job.get("redirect_url", "")
            if url in existing_urls:
                skipped += 1
                continue
            existing_urls.add(url)
            fill = odd_fill if (next_row % 2) else even_fill
            desc = job.get("description", "").replace("\n", " ")[:300]
            row_data = [
                job.get("title", ""),
                job.get("company", {}).get("display_name", ""),
                job.get("location", {}).get("display_name", ""),
                job.get("salary_min") or "",
                job.get("salary_max") or "",
                job.get("created", "")[:10],
                url, desc, now_str,
            ]
            for col, val in enumerate(row_data, 1):
                cell = ws.cell(row=next_row, column=col, value=val)
                cell.fill = fill
                cell.alignment = Alignment(vertical="center", wrap_text=(col == 8))
                if col == 7 and url:
                    cell.font = link_font; cell.hyperlink = url
                else:
                    cell.font = data_font
            ws.row_dimensions[next_row].height = 18
            next_row += 1; added += 1

        for col, w in enumerate([40, 25, 22, 16, 16, 14, 40, 60, 18], 1):
            ws.column_dimensions[ws.cell(1, col).column_letter].width = w

        if "Search Info" in wb.sheetnames:
            del wb["Search Info"]
        meta    = wb.create_sheet("Search Info")
        bg_fill = PatternFill("solid", start_color="1e293b")
        m_font  = Font(name="Arial", size=10, color="F1F5F9")
        m_bold  = Font(name="Arial", size=10, color="38bdf8", bold=True)
        total   = ws.max_row - 1

        for r, (label, value) in enumerate([
            ("Search query",       self.what_var.get()),
            ("Location",           self.where_var.get()),
            ("Country",            self.country_var.get()),
            ("Last updated",       now_str),
            ("Total jobs",         total),
            ("Added this run",     added),
            ("Duplicates skipped", skipped),
            ("Source",             "Adzuna API — adzuna.com"),
        ], 1):
            c1 = meta.cell(row=r, column=1, value=label)
            c2 = meta.cell(row=r, column=2, value=value)
            c1.font = m_bold; c1.fill = bg_fill
            c2.font = m_font; c2.fill = bg_fill

        meta.column_dimensions["A"].width = 22
        meta.column_dimensions["B"].width = 35
        meta.sheet_view.showGridLines = False
        ws.sheet_view.showGridLines   = False

        wb.save(path)
        self._set_status(
            f"✓ {added} new  |  {skipped} duplicates skipped  |  {total} total in {os.path.basename(path)}")
        if not silent and sys.platform == "win32":
            subprocess.Popen(f'explorer /select,"{os.path.abspath(path)}"')

    # ── Schedule Task ─────────────────────────────────────────────────────────
    def _schedule(self):
        if sys.platform != "win32":
            messagebox.showinfo("Not supported", "Task Scheduler is only available on Windows.")
            return

        time_str = self.sched_time_var.get().strip()
        try:
            hour, minute = time_str.split(":")
            int(hour); int(minute)
        except Exception:
            messagebox.showerror("Invalid time", "Please enter time as HH:MM (e.g. 08:00)")
            return

        # Save current search params so job_search_auto.py uses them
        _save_config({
            "searches": [{
                "what":    self.what_var.get().strip(),
                "where":   self.where_var.get().strip(),
                "country": COUNTRIES.get(self.country_var.get(), "de"),
            }]
        })

        python_exe  = sys.executable
        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "job_search_auto.py")
        task_name   = "JobSearchDaily"

        if not os.path.exists(script_path):
            messagebox.showerror("Missing file",
                                 f"job_search_auto.py not found:\n{script_path}\n\n"
                                 "Please place job_search_auto.py in the same folder.")
            return

        cmd = (f'schtasks /create /f /tn "{task_name}" '
               f'/tr "\\"{python_exe}\\" \\"{script_path}\\"" '
               f'/sc daily /st {hour.zfill(2)}:{minute.zfill(2)} '
               f'/ru "{os.environ.get("USERNAME", "")}" /rl limited')

        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            self._set_status(f"✓ Scheduled daily at {hour.zfill(2)}:{minute.zfill(2)}")
            messagebox.showinfo("✓ Scheduled",
                f"Daily auto-run set for {hour.zfill(2)}:{minute.zfill(2)}\n\n"
                f"Task name: {task_name}\nScript: job_search_auto.py\n\n"
                "To remove: click '✕ Remove Schedule'")
        else:
            # Fallback without /ru
            cmd2 = (f'schtasks /create /f /tn "{task_name}" '
                    f'/tr "\\"{python_exe}\\" \\"{script_path}\\"" '
                    f'/sc daily /st {hour.zfill(2)}:{minute.zfill(2)}')
            result2 = subprocess.run(cmd2, shell=True, capture_output=True, text=True)
            if result2.returncode == 0:
                self._set_status(f"✓ Scheduled daily at {hour.zfill(2)}:{minute.zfill(2)}")
                messagebox.showinfo("✓ Scheduled",
                    f"Daily auto-run set for {hour.zfill(2)}:{minute.zfill(2)}\n\n"
                    "To remove: click '✕ Remove Schedule'")
            else:
                self._set_status("⚠ Scheduling failed — try Run as Administrator", ok=False)
                messagebox.showerror("Access Denied",
                    "Windows blocked the scheduler.\n\n"
                    "Fix: Right-click job_search_3.py → 'Run as administrator'\n"
                    "then click ⏱ Schedule Daily again.")

    # ── Remove Scheduled Task ─────────────────────────────────────────────────
    def _unschedule(self):
        if sys.platform != "win32":
            messagebox.showinfo("Not supported", "Task Scheduler is only available on Windows.")
            return

        task_name = "JobSearchDaily"
        confirm = messagebox.askyesno(
            "Remove Schedule",
            f"Remove the daily auto-run task '{task_name}'?\n\n"
            "The app and Excel files are not affected.")
        if not confirm:
            return

        result = subprocess.run(
            f'schtasks /delete /f /tn "{task_name}"',
            shell=True, capture_output=True, text=True)

        if result.returncode == 0:
            self._set_status("✓ Daily schedule removed")
            messagebox.showinfo("✓ Removed",
                f"Task '{task_name}' has been removed.\n\n"
                "You can re-schedule anytime with ⏱ Schedule Daily.")
        else:
            err = result.stderr.strip().lower()
            if "cannot find" in err or "nicht gefunden" in err:
                self._set_status("⚠ No active schedule found", ok=False)
                messagebox.showwarning("Not found",
                    f"No task named '{task_name}' was found.\n"
                    "It may have already been removed.")
            else:
                self._set_status("⚠ Could not remove schedule", ok=False)
                messagebox.showerror("Error", f"schtasks error:\n{result.stderr}")


if __name__ == "__main__":
    try:
        JobSearchApp().mainloop()
    except Exception as e:
        traceback.print_exc()
        input("Press Enter to close...")
