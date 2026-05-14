import requests
from .base import BaseProvider, AuthError


class AdzunaProvider(BaseProvider):
    name = "Adzuna"
    _URL = "https://api.adzuna.com/v1/api/jobs/{country}/search/1"

    def __init__(self, app_id: str, app_key: str):
        self.app_id  = app_id
        self.app_key = app_key

    _PAGE_SIZE = 50  # Adzuna API hard limit per page
    _URL = "https://api.adzuna.com/v1/api/jobs/{country}/search/{page}"

    def search(self, what, where, country="de", results=20,
               sort_by="date", salary_min=None, salary_max=None,
               full_time=False, permanent=False, **kwargs):
        base_params = {
            "app_id":           self.app_id,
            "app_key":          self.app_key,
            "what":             what,
            "where":            where,
            "results_per_page": self._PAGE_SIZE,
            "sort_by":          sort_by,
        }
        if salary_min: base_params["salary_min"] = salary_min
        if salary_max: base_params["salary_max"] = salary_max
        if full_time:  base_params["full_time"]  = 1
        if permanent:  base_params["permanent"]  = 1

        all_jobs = []
        page = 1
        try:
            while len(all_jobs) < results:
                url = self._URL.format(country=country, page=page)
                resp = requests.get(url, params=base_params, timeout=10)
                if resp.status_code in (401, 403):
                    raise AuthError("Adzuna")
                if resp.status_code != 200:
                    break
                batch = resp.json().get("results", [])
                if not batch:
                    break
                all_jobs.extend(self._normalize(j) for j in batch)
                if len(batch) < self._PAGE_SIZE:
                    break
                page += 1
        except AuthError:
            raise
        except Exception:
            pass
        return all_jobs[:results]

    def _normalize(self, j):
        return self._norm(
            title       = j.get("title", ""),
            company     = j.get("company", {}).get("display_name", ""),
            location    = j.get("location", {}).get("display_name", ""),
            salary_min  = j.get("salary_min") or "",
            salary_max  = j.get("salary_max") or "",
            created     = j.get("created", ""),
            url         = j.get("redirect_url", ""),
            description = j.get("description", ""),
        )
