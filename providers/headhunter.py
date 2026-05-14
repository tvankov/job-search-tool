import requests
from .base import BaseProvider, AuthError

_URL = "https://api.hh.ru/vacancies"
_HEADERS = {"User-Agent": "JobSearchTool/1.0 (support@jobsearchtool.com)"}
_PER_PAGE = 100


class HeadHunterProvider(BaseProvider):
    name = "HeadHunter"

    def __init__(self, token=""):
        self.token = token.strip()

    def search(self, what, where, country="ru", results=20, **kwargs):
        if not self.token:
            return []
        headers = {**_HEADERS, "Authorization": f"Bearer {self.token}"}
        try:
            all_jobs = []
            page     = 0
            while len(all_jobs) < results:
                params = {"area": 113, "per_page": _PER_PAGE, "page": page}
                if what.strip():
                    params["text"] = what.strip()
                resp = requests.get(_URL, params=params, headers=headers, timeout=10)
                if resp.status_code in (401, 403):
                    raise AuthError("HeadHunter")
                if resp.status_code != 200:
                    break
                items = resp.json().get("items", [])
                if not items:
                    break
                all_jobs.extend(self._normalize(j) for j in items)
                if len(items) < _PER_PAGE:
                    break
                page += 1
            return all_jobs[:results]
        except AuthError:
            raise
        except Exception:
            pass
        return []

    def _normalize(self, j):
        salary   = j.get("salary") or {}
        s_from   = salary.get("from") or ""
        s_to     = salary.get("to")   or ""
        currency = salary.get("currency", "")
        snippet  = j.get("snippet") or {}
        desc     = " ".join(filter(None, [
            snippet.get("responsibility", ""),
            snippet.get("requirement", ""),
        ]))
        sal_min = f"{s_from} {currency}".strip() if s_from else ""
        sal_max = f"{s_to} {currency}".strip()   if s_to   else ""
        return self._norm(
            title       = j.get("name", ""),
            company     = j.get("employer", {}).get("name", ""),
            location    = j.get("area", {}).get("name", ""),
            salary_min  = sal_min,
            salary_max  = sal_max,
            url         = j.get("alternate_url", ""),
            created     = j.get("published_at", ""),
            description = desc,
        )
