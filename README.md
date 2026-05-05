# Job Search Tool

A desktop app for searching and automatically collecting job listings — built with Python and the Adzuna API.

![Python](https://img.shields.io/badge/Python-3.8+-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## Features

- **Live job search** — search by job title, location, country, salary range and more
- **19 countries supported** — Germany, USA, UK, Austria, Switzerland, France, Australia and more
- **Auto save to Excel** — results are saved automatically on first search, duplicates skipped
- **Saved Results browser** — view and manage all your saved Excel files inside the app
- **Daily Auto Run** — schedule the scraper to run every day via Windows Task Scheduler
- **Search history** — see the last 30 auto-runs with date, time and job counts
- **Persistent credentials** — your API keys are saved locally in `config.json`
- **Dark UI** — clean, modern dark theme built with Tkinter

---

## Requirements

- Python 3.8 or higher
- A free Adzuna API account → [developer.adzuna.com](https://developer.adzuna.com)

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/YOUR_USERNAME/job-search-tool.git
cd job-search-tool

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
python job_search_4.py
```

---

## Getting your free API credentials

1. Go to [developer.adzuna.com](https://developer.adzuna.com) and register
2. Create a new application (any name)
3. Copy your **App ID** and **App Key**
4. Open the app → go to the **Credentials** tab → paste them in and click **Save Credentials**

---

## Usage

| Tab | Description |
|---|---|
| **Search** | Search for jobs live. Double-click a row to open the listing. |
| **Saved Results** | Browse all saved Excel files. Click to view contents, delete if not needed. |
| **Auto Run** | Schedule a daily search. Uses the last Job Title and Location you searched. |
| **Credentials** | Enter and save your Adzuna API keys. |
| **Help** | Step-by-step guide and FAQ. |
| **About** | App info and links. |

---

## Project structure

```
job-search-tool/
├── job_search_4.py      # Main GUI app
├── job_search_auto.py   # Headless runner for Task Scheduler
├── requirements.txt
├── README.md
└── .gitignore
```

> `config.json` and the `jobs/` folder are created automatically on first use and are excluded from version control.

---

## Built with

- [Python](https://python.org) + Tkinter
- [openpyxl](https://openpyxl.readthedocs.io) — Excel export
- [Adzuna API](https://developer.adzuna.com) — job data

---

## Author

**Todor Vankov**  
[www.todorvankov.com](https://www.todorvankov.com)

---

## License

MIT — free to use and modify.
