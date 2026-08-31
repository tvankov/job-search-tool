from .base import BaseProvider, AuthError, logger


class JoobleProvider(BaseProvider):
    name = "Jooble"
    _URL = "https://jooble.org/api/{key}"

    def __init__(self, api_key: str):
        self.api_key = api_key

    _PAGE_SIZE = 100  # Jooble max per page

    def search(self, what, where, country="de", results=20, **kwargs):
        if not self.api_key:
            return []
        all_jobs = []
        page = 1
        try:
            while len(all_jobs) < results:
                body = {"resultonpage": self._PAGE_SIZE, "page": page}
                if what.strip():
                    body["keywords"] = what.strip()
                if where.strip():
                    body["location"] = where.strip()
                resp = self._request(
                    "post",
                    self._URL.format(key=self.api_key),
                    json=body,
                )
                if resp is None:
                    break
                if resp.status_code in (401, 403):
                    raise AuthError("Jooble")
                if resp.status_code != 200:
                    break
                batch = resp.json().get("jobs", [])
                if not batch:
                    break
                all_jobs.extend(self._normalize(j) for j in batch)
                if len(batch) < self._PAGE_SIZE:
                    break
                page += 1
        except AuthError:
            raise
        except ValueError as exc:  # malformed JSON response
            logger.warning("%s: bad JSON — %s", self.name, exc)
        return all_jobs[:results]

    def _normalize(self, j):
        return self._norm(
            title       = j.get("title", ""),
            company     = j.get("company", ""),
            location    = j.get("location", ""),
            salary_min  = j.get("salary", ""),
            url         = j.get("link", ""),
            created     = j.get("updated", ""),
            description = j.get("snippet", ""),
        )
