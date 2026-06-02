"""
job_search_auto.py
──────────────────
Headless version — runs without GUI, saves results to Excel.
Designed to be called by Windows Task Scheduler once a day.
"""

import os
import json
import requests
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Load config from config.json (written by the GUI) ────────────────────────
_BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
_CONFIG_PATH = os.path.join(_BASE_DIR, "config.json")

def _load_config():
    if os.path.exists(_CONFIG_PATH):
        try:
            with open(_CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

_cfg = _load_config()

APP_ID           = _cfg.get("app_id",  "")
APP_KEY          = _cfg.get("app_key", "")
SEARCHES         = _cfg.get("searches", [])
RESULTS_PER_PAGE = 50
# ─────────────────────────────────────────────────────────────────────────────

BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jobs")
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "job_search.log")


def log(msg: str):
    """Write message to log file and print to console."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def fetch_jobs(what: str, where: str, country: str) -> list:
    """Fetch jobs from Adzuna API."""
    url = f"https://api.adzuna.com/v1/api/jobs/{country}/search/1"
    params = {
        "app_id":           APP_ID,
        "app_key":          APP_KEY,
        "what":             what,
        "where":            where,
        "results_per_page": RESULTS_PER_PAGE,
        "sort_by":          "date",
    }
    try:
        resp = requests.get(url, params=params, timeout=15)
        if resp.status_code == 200:
            return resp.json().get("results", [])
        else:
            log(f"  !! API error {resp.status_code} for '{what}' in '{where}'")
            return []
    except Exception as e:
        log(f"  !! Connection error: {e}")
        return []


def save_to_excel(jobs: list, job_title: str, search_info: dict):
    """Append new jobs to the persistent Excel file for this job title + location."""
    where = search_info.get("where", "").strip()
    label = f"{job_title} {where}".strip() if where else job_title
    safe  = label.replace(" ", "_")
    folder    = os.path.join(BASE_DIR, safe)
    os.makedirs(folder, exist_ok=True)
    path      = os.path.join(folder, f"{safe}.xlsx")

    headers   = ["Title", "Company", "Location",
                 "Salary Min (€)", "Salary Max (€)", "Posted",
                 "Link", "Description", "Imported On"]

    hdr_fill  = PatternFill("solid", start_color="0ea5e9")
    hdr_font  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    hdr_align = Alignment(horizontal="center", vertical="center")
    border    = Border(bottom=Side(style="thin", color="334155"))
    odd_fill  = PatternFill("solid", start_color="1e293b")
    even_fill = PatternFill("solid", start_color="162032")
    data_font = Font(name="Arial", size=10, color="F1F5F9")
    link_font = Font(name="Arial", size=10, color="38bdf8", underline="single")

    # Load existing or create new
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb["Job Results"]
        existing_urls = {str(row[6]) for row in ws.iter_rows(min_row=2, values_only=True) if row[6]}
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Job Results"
        existing_urls = set()
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = hdr_font; cell.fill = hdr_fill
            cell.alignment = hdr_align; cell.border = border
        ws.row_dimensions[1].height = 22

    # Append new jobs
    next_row = ws.max_row + 1
    added = skipped = 0
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    for job in jobs:
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
            url,
            desc,
            now_str,
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=next_row, column=col, value=val)
            cell.fill = fill
            cell.alignment = Alignment(vertical="center", wrap_text=(col == 8))
            cell.font = (link_font if col == 7 and url else data_font)
            if col == 7 and url:
                cell.hyperlink = url
        ws.row_dimensions[next_row].height = 18
        next_row += 1
        added += 1

    # Column widths
    for col, w in enumerate([40, 25, 22, 16, 16, 14, 40, 60, 18], 1):
        ws.column_dimensions[ws.cell(1, col).column_letter].width = w

    # Update Search Info sheet
    if "Search Info" in wb.sheetnames:
        del wb["Search Info"]
    meta      = wb.create_sheet("Search Info")
    bg        = PatternFill("solid", start_color="1e293b")
    m_font    = Font(name="Arial", size=10, color="F1F5F9")
    m_bold    = Font(name="Arial", size=10, color="38bdf8", bold=True)
    total     = ws.max_row - 1

    for r, (label, value) in enumerate([
        ("Search query",       search_info.get("what", "")),
        ("Location",           search_info.get("where", "")),
        ("Last updated",       now_str),
        ("Total jobs",         total),
        ("Added this run",     added),
        ("Duplicates skipped", skipped),
        ("Source",             "Adzuna API — adzuna.com"),
    ], 1):
        c1 = meta.cell(row=r, column=1, value=label)
        c2 = meta.cell(row=r, column=2, value=value)
        c1.font = m_bold;  c1.fill = bg
        c2.font = m_font;  c2.fill = bg

    meta.column_dimensions["A"].width = 22
    meta.column_dimensions["B"].width = 35
    meta.sheet_view.showGridLines = False
    ws.sheet_view.showGridLines   = False

    wb.save(path)
    return added, skipped, total


def run():
    log("=" * 55)
    log("Job Search Auto-Run started")
    log("=" * 55)

    if not APP_ID or not APP_KEY:
        log("!! No API credentials found. Open the app -> Credentials tab -> Save Credentials.")
        log("=" * 55 + "\n")
        return

    if not SEARCHES:
        log("!! No search configured. Open the app -> Search tab -> click Schedule Daily.")
        log("=" * 55 + "\n")
        return

    total_added = 0

    for search in SEARCHES:
        what, where, country = search["what"], search["where"], search["country"]
        log(f"Searching: '{what}' in '{where}' ({country.upper()})")

        jobs = fetch_jobs(what, where, country)
        log(f"  >> {len(jobs)} jobs fetched from API")

        if jobs:
            added, skipped, total = save_to_excel(jobs, what, search)
            log(f"  >> {added} new  |  {skipped} duplicates skipped  |  {total} total in file")
            total_added += added

    log(f"Done — {total_added} new jobs added across all searches")
    log("=" * 55 + "\n")


if __name__ == "__main__":
    run()
