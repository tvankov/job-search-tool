import requests
from .base import BaseProvider


class RemoteOKProvider(BaseProvider):
    name = "RemoteOK"
    _URL = "https://remoteok.com/api"

    def search(self, what, where, country="de", results=20, **kwargs):
        try:
            resp = requests.get(
                self._URL,
                headers={"User-Agent": "JobSearchTool/1.0"},
                timeout=15,
            )
            print(f"[RemoteOK] HTTP {resp.status_code}")
            if resp.status_code == 200:
                data      = resp.json()
                what_low  = what.lower()
                words     = what_low.split() if what_low else []
                jobs_raw  = [j for j in data if isinstance(j, dict) and "position" in j]
                print(f"[RemoteOK] total jobs in feed: {len(jobs_raw)}")
                if words:
                    def _rok_match(j):
                        haystack = (j.get("position", "") + " " +
                                    " ".join(j.get("tags", []))).lower()
                        return all(w in haystack for w in words)
                    jobs_raw = [j for j in jobs_raw if _rok_match(j)]
                    print(f"[RemoteOK] after keyword filter '{what_low}': {len(jobs_raw)}")
                return [self._normalize(j) for j in jobs_raw]
        except Exception as e:
            print(f"[RemoteOK] Error: {e}")
        return []

    def _normalize(self, j):
        return self._norm(
            title       = j.get("position", ""),
            company     = j.get("company", ""),
            location    = j.get("location", "") or "Remote",
            salary_min  = j.get("salary_min") or "",
            salary_max  = j.get("salary_max") or "",
            url         = j.get("url", ""),
            created     = j.get("date", ""),
            description = j.get("description", ""),
        )
