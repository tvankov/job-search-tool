import time
import requests
from .base import BaseProvider


class HimalayasProvider(BaseProvider):
    name = "Himalayas"
    _SEARCH_URL = "https://himalayas.app/jobs/api/search"
    _BROWSE_URL = "https://himalayas.app/jobs/api"

    def search(self, what, where, country="", results=20, **kwargs):
        all_jobs = []
        offset   = 0
        limit    = 20
        try:
            while len(all_jobs) < results:
                params = {"limit": limit, "offset": offset}
                if what.strip():
                    params["q"] = what.strip()
                if where.strip():
                    params["country"] = where.strip()

                url  = self._SEARCH_URL if what.strip() else self._BROWSE_URL
                resp = requests.get(url, params=params, timeout=10)
                if resp.status_code == 429:
                    time.sleep(2)
                    continue
                if resp.status_code != 200:
                    break
                data       = resp.json()
                batch      = data.get("jobs", [])
                total      = data.get("totalCount", 0)
                if not batch:
                    break
                all_jobs.extend(self._normalize(j) for j in batch)
                offset += limit
                if offset >= total:
                    break
        except Exception:
            pass
        return all_jobs[:results]

    def _normalize(self, j):
        # location: list like ["Germany", "USA"] or empty
        loc = j.get("locationRestrictions") or []
        if isinstance(loc, list):
            location = ", ".join(loc) if loc else "Worldwide Remote"
        else:
            location = str(loc) or "Worldwide Remote"

        # salary
        sal_min = j.get("minSalary") or ""
        sal_max = j.get("maxSalary") or ""
        if sal_min == "None": sal_min = ""
        if sal_max == "None": sal_max = ""

        # date: unix timestamp → readable
        pub = j.get("pubDate", "")
        if pub:
            try:
                from datetime import datetime
                pub = datetime.utcfromtimestamp(int(pub)).strftime("%Y-%m-%d")
            except Exception:
                pub = str(pub)

        return self._norm(
            title       = j.get("title", ""),
            company     = j.get("companyName", ""),
            location    = location,
            salary_min  = sal_min,
            salary_max  = sal_max,
            created     = pub,
            url         = j.get("applicationLink", j.get("guid", "")),
            description = j.get("excerpt", ""),
        )
