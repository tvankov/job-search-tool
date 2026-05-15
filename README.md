# Job Search Tool

A desktop app for searching, collecting and analysing job listings — built with Python and Tkinter.

![Version](https://img.shields.io/badge/Version-1.0.1-blue)
![Python](https://img.shields.io/badge/Python-3.8+-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

<div align="center">
<img width="500" alt="grafik" src="https://github.com/user-attachments/assets/2ff96dfb-173f-427a-92ba-a6c22a1ebb48" />
</div>

---

## Features

- **12 job providers** — search across Adzuna, Reed, Findwork, Jooble, HeadHunter, Arbeitnow, Bundesagentur, The Muse, RemoteOK, WeWorkRemotely, Remotive and Himalayas simultaneously
- **Provider chips** — enable/disable providers with one click; selection is saved between sessions
- **Live job search** — filter by job title, location, country, salary range, remote-only and more
- **Auto-enable on save** — entering an API key and clicking Save Credentials automatically activates that provider
- **Auto save to Excel** — results saved on first search, duplicates skipped
- **Saved Results browser** — view and manage all saved Excel files inside the app
- **Analytics tab** — charts and statistics over your saved results
- **Daily Auto Run** — schedule the scraper via Windows Task Scheduler; runs even on battery
- **Update check** — the app checks GitHub for new versions on startup
- **Bug reporting** — errors are logged to `error.log`; Report Bug dialog opens it directly
- **Fully dark UI** — consistent dark theme across all controls (tables, dropdowns, scrollbars)

---

| | | |
|---|---|---|
| <img width="400" src="https://github.com/user-attachments/assets/c0bfcce2-531d-4bb2-ae16-18e13419fb2b" /> | <img width="400" src="https://github.com/user-attachments/assets/9e5075f9-8860-4e25-bd0b-59e3e180638c" /> | <img width="400" src="https://github.com/user-attachments/assets/8e617389-0e59-4272-917d-fd674e54f231" /> |

---

## Supported providers

| Provider | Key required | Coverage |
|---|---|---|
| **Adzuna** | Yes (free) | 19 countries |
| **Reed** | Yes (free) | UK |
| **Findwork** | Yes (free) | Tech / Remote |
| **Jooble** | Yes (free) | 70+ countries |
| **HeadHunter** | Yes (free) | Russia / CIS |
| **Bundesagentur** | No | Germany |
| **Arbeitnow** | No | Europe |
| **The Muse** | No | USA |
| **RemoteOK** | No | Remote |
| **WeWorkRemotely** | No | Remote |
| **Remotive** | No | Remote |
| **Himalayas** | No | Remote |

> RemoteOK and WeWorkRemotely use public RSS feeds — for personal use only.

---

## Installation

### Option A — Windows Installer (recommended)

Download `JobSearchTool_Setup_v1.0.1.exe` from the [Releases page](https://github.com/tvankov/job-search-tool/releases/latest) and run it. No Python required.

### Option B — Run from source

```bash
# 1. Clone the repository
git clone https://github.com/tvankov/job-search-tool.git
cd job-search-tool

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
python job_search_4.py
```

---

## Getting API keys

All keys are free. Enter them in the **Credentials** tab — the provider chip turns green automatically when a key is saved.

| Provider | Where to get the key |
|---|---|
| **Adzuna** | [developer.adzuna.com](https://developer.adzuna.com) — register, create an app, copy App ID + App Key |
| **Reed** | [reed.co.uk/developers](https://www.reed.co.uk/developers/jobseeker) — register, copy API Key |
| **Findwork** | [findwork.dev/api-token-auth](https://findwork.dev/api-token-auth/) — register, copy token |
| **Jooble** | [jooble.org/api](https://jooble.org/api/about) — request a key by email |
| **HeadHunter** | [dev.hh.ru](https://dev.hh.ru) — create an app, complete OAuth flow, copy access token |

---

## Usage

| Tab | Description |
|---|---|
| **Search** | Search jobs live. Toggle provider chips on/off. Double-click a row to open the listing. |
| **Saved Results** | Browse all saved Excel files. Click a file to view its jobs. |
| **Auto Run** | Schedule a daily search and view run history. |
| **Analytics** | Charts and statistics over your saved results. |
| **Credentials** | Enter API keys. Saving a key automatically enables that provider. |
| **Help** | Step-by-step guide and FAQ. |
| **About** | App version and links. |

---

## Project structure

```
job-search-tool/
├── job_search_4.py        # Main GUI app
├── job_search_auto.py     # Headless runner for Task Scheduler
├── version.json           # Version info for update check
├── providers/             # One module per job provider
│   ├── base.py
│   ├── adzuna.py
│   ├── reed.py
│   ├── findwork.py
│   ├── jooble.py
│   ├── headhunter.py
│   ├── arbeitnow.py
│   ├── bundesagentur.py
│   ├── themuse.py
│   ├── remoteok.py
│   ├── weworkremotely.py
│   ├── remotive.py
│   └── himalayas.py
├── requirements.txt
├── installer.iss          # Inno Setup script
├── README.md
└── .gitignore
```

> `config.json`, the `jobs/` folder and `error.log` are created automatically on first use and are excluded from version control.

---

## Built with

- [Python](https://python.org) + Tkinter — GUI
- [requests](https://docs.python-requests.org) — API calls
- [openpyxl](https://openpyxl.readthedocs.io) — Excel export
- [matplotlib](https://matplotlib.org) — Analytics charts
- [lxml](https://lxml.de) — RSS feed parsing
- [PyInstaller](https://pyinstaller.org) — Windows executable
- [Inno Setup](https://jrsoftware.org/isinfo.php) — Windows installer

---

## Author

**Todor Vankov**  
[www.todorvankov.com](https://www.todorvankov.com)

---

## License

MIT — free to use and modify.
