import requests
from .base import BaseProvider, AuthError


class ReedProvider(BaseProvider):
    name = "Reed"
    _URL = "https://www.reed.co.uk/api/1.0/search"

    def __init__(self, api_key: str):
        self.api_key = api_key

    _PAGE_SIZE = 100  # Reed max per page

    def search(self, what, where, country="gb", results=20, **kwargs):
        if not self.api_key:
            return []
        all_jobs = []
        skip = 0
        try:
            while len(all_jobs) < results:
                resp = requests.get(
                    self._URL,
                    params={"keywords": what, "locationName": where,
                            "resultsToTake": self._PAGE_SIZE,
                            "resultsToSkip": skip},
                    auth=(self.api_key, ""),
                    timeout=10,
                )
                if resp.status_code in (401, 403):
                    raise AuthError("Reed")
                if resp.status_code != 200:
                    break
                batch = resp.json().get("results", [])
                if not batch:
                    break
                all_jobs.extend(self._normalize(j) for j in batch)
                if len(batch) < self._PAGE_SIZE:
                    break
                skip += self._PAGE_SIZE
        except AuthError:
            raise
        except Exception:
            pass
        return all_jobs[:results]

    def _normalize(self, j):
        return self._norm(
            title       = j.get("jobTitle", ""),
            company     = j.get("employerName", ""),
            location    = j.get("locationName", ""),
            salary_min  = j.get("minimumSalary") or "",
            salary_max  = j.get("maximumSalary") or "",
            created     = j.get("date", ""),
            url         = j.get("jobUrl", ""),
            description = j.get("jobDescription", ""),
        )
