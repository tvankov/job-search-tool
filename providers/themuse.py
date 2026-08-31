from .base import BaseProvider, logger


class TheMuseProvider(BaseProvider):
    name = "The Muse"
    _URL = "https://www.themuse.com/api/public/jobs"

    def __init__(self, api_key: str = ""):
        self.api_key = api_key

    def search(self, what, where, country="us", results=20, **kwargs):
        try:
            what_low  = what.lower()
            words     = what_low.split() if what_low else []
            collected = []
            page      = 1
            while len(collected) < results:
                params = {"page": page, "descending": "true"}
                if self.api_key:
                    params["api_key"] = self.api_key
                resp = self._request("get", self._URL, params=params)
                if resp is None:
                    break
                if resp.status_code != 200:
                    break
                jobs = resp.json().get("results", [])
                if not jobs:
                    break
                for j in jobs:
                    if words:
                        haystack = j.get("name", "").lower()
                        if not all(w in haystack for w in words):
                            continue
                    collected.append(self._normalize(j))
                page += 1
            return collected[:results]
        except ValueError as exc:  # malformed JSON response
            logger.warning("%s: bad JSON — %s", self.name, exc)
        return []

    def _normalize(self, j):
        locs = j.get("locations", [])
        return self._norm(
            title    = j.get("name", ""),
            company  = j.get("company", {}).get("name", ""),
            location = locs[0]["name"] if locs else "USA",
            url      = j.get("refs", {}).get("landing_page", ""),
            created  = j.get("publication_date", ""),
        )
