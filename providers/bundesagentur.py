import requests
from .base import BaseProvider

_HEADERS = {
    "X-API-Key": "jobboerse-jobsuche",
    "User-Agent": "JobSearchTool/1.0",
}
_URL = "https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v4/jobs"
_PAGE_SIZE = 100


class BundesagenturProvider(BaseProvider):
    name = "Bundesagentur"

    def search(self, what, where, country="de", results=20, **kwargs):
        try:
            all_jobs = []
            page     = 1
            while len(all_jobs) < results:
                params = {"page": page, "size": _PAGE_SIZE}
                if what.strip(): params["was"] = what.strip()
                if where.strip(): params["wo"]  = where.strip()
                resp = requests.get(_URL, params=params, headers=_HEADERS, timeout=10)
                if resp.status_code != 200:
                    break
                batch = resp.json().get("stellenangebote", [])
                if not batch:
                    break
                all_jobs.extend(self._normalize(j) for j in batch)
                if len(batch) < _PAGE_SIZE:
                    break
                page += 1
            return all_jobs[:results]
        except Exception:
            pass
        return []

    def _normalize(self, j):
        ort   = j.get("arbeitsort", {})
        refnr = j.get("refnr", "")
        url   = j.get("externeUrl") or (
            f"https://www.arbeitsagentur.de/jobsuche/jobdetail/{refnr}" if refnr else ""
        )
        return self._norm(
            title    = j.get("titel", ""),
            company  = j.get("arbeitgeber", ""),
            location = ", ".join(filter(None, [ort.get("ort", ""), ort.get("land", "")])),
            url      = url,
            created  = j.get("aktuelleVeroeffentlichungsdatum", ""),
        )
