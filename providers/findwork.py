import requests
from .base import BaseProvider, AuthError


class FindworkProvider(BaseProvider):
    name = "Findwork"
    _URL = "https://findwork.dev/api/jobs/"

    def __init__(self, api_key: str):
        self.api_key = api_key

    def search(self, what, where, country="", results=20, **kwargs):
        if not self.api_key:
            return []
        headers = {"Authorization": f"Token {self.api_key}"}
        params = {
            "search":  what,
            "sort_by": "date",
            "limit":   50,
        }
        if where:
            params["location"] = where

        jobs = []
        url = self._URL
        try:
            while url and len(jobs) < results:
                resp = requests.get(url, params=params, headers=headers, timeout=10)
                if resp.status_code in (401, 403):
                    raise AuthError("Findwork")
                if resp.status_code != 200:
                    break
                data = resp.json()
                jobs += [self._normalize(j) for j in data.get("results", [])]
                url = data.get("next")   # next page URL or None
                params = {}              # next URL already has all params encoded
        except AuthError:
            raise
        except Exception:
            pass
        return jobs[:results]

    def _normalize(self, j):
        return self._norm(
            title       = j.get("role", ""),
            company     = j.get("company_name", ""),
            location    = j.get("location", "Remote"),
            url         = j.get("url", ""),
            created     = j.get("date_posted", ""),
            description = j.get("text", ""),
        )
