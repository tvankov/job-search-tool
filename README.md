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
| <img width="1223" alt="grafik" src="https://github.com/user-attachments/assets/a3732027-529c-46b6-8bc5-39a313882a14" /> | <img width="1225" alt="grafik" src="https://github.com/user-attachments/assets/2dd47599-7cef-49f3-ae68-799d6dbf8225" /> | <img width="1222" alt="grafik" src="https://github.com/user-attachments/assets/d502724b-6cc7-45af-be6d-db96f20eba95" /> |

| | | |
|---|---|---|
| <img width="400" alt="grafik" src="https://github.com/user-attachments/assets/edceb730-e3d7-4d17-b99e-d3671aa91e27" /> | <img width="400" alt="grafik" src="https://github.com/user-attachments/assets/83f718ef-6e4c-4e11-8c9f-3a1bea400f51" /> | <img width="400" alt="grafik" src="https://github.com/user-attachments/assets/4b12e5e5-6217-4c8f-8e1f-2809d4d614a3" /> |

| | | |
|---|---|---|
| <img width="400" alt="grafik" src="https://github.com/user-attachments/assets/2a7aafaf-be72-443d-b319-3b1c65fdaf1b" /> | <img width="400"  alt="grafik" src="https://github.com/user-attachments/assets/8fb4064d-7896-4171-b007-044f30588dbc" /> | <img width="400"  alt="grafik" src="https://github.com/user-attachments/assets/880ea6df-1b0c-4a1c-9fe5-2a1d9abf6a66" />|


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
