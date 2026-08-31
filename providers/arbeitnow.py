from .base import BaseProvider, logger


class ArbeitnowProvider(BaseProvider):
    name = "Arbeitnow"
    _URL = "https://www.arbeitnow.com/api/job-board-api"

    def search(self, what, where, country="de", results=20, **kwargs):
        try:
            collected = []
            what_low  = what.lower()
            where_low = where.lower()
            words     = what_low.split() if what_low else []
            for page in range(1, 51):  # up to 50 pages
                resp = self._request("get", self._URL, params={"page": page})
                if resp is None:
                    break
                if resp.status_code != 200:
                    break
                data = resp.json().get("data", [])
                if not data:
                    break
                for j in data:
                    if words:
                        haystack = (j.get("title", "") + " " +
                                    j.get("description", "")).lower()
                        if not all(w in haystack for w in words):
                            continue
                    if where_low and where_low not in j.get("location", "").lower():
                        continue
                    collected.append(self._normalize(j))
                if len(collected) >= results:
                    break
            return collected[:results]
        except ValueError as exc:  # malformed JSON response
            logger.warning("%s: bad JSON — %s", self.name, exc)
        return []

    def _normalize(self, j):
        return self._norm(
            title       = j.get("title", ""),
            company     = j.get("company_name", ""),
            location    = j.get("location", ""),
            url         = j.get("url", ""),
            created     = j.get("created_at", ""),
            description = j.get("description", ""),
        )
