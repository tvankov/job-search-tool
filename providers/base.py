import time
import logging
from abc import ABC, abstractmethod

import requests

logger = logging.getLogger("jobsearch.providers")


class AuthError(Exception):
    """Raised when a provider rejects the API key (HTTP 401 / 403)."""


class BaseProvider(ABC):
    name: str = ""

    # HTTP statuses worth a retry: rate-limit + transient server errors.
    _RETRY_STATUS = frozenset({429, 500, 502, 503, 504})
    _MAX_RETRIES  = 1
    _BACKOFF_SECS = 1.5

    @abstractmethod
    def search(self, what: str, where: str, country: str,
               results: int, **kwargs) -> list:
        """Return list of normalized job dicts."""

    def _request(self, method: str, url: str, **kwargs) -> "requests.Response | None":
        """HTTP request with a default timeout and one retry on rate-limit /
        server errors. Returns the Response, or None if the request failed at
        the network level (timeout / connection error) — already logged.

        Callers still handle 401/403 (AuthError) and non-200 responses
        themselves; this only centralizes timeouts, retries and network-error
        logging so no provider swallows failures silently.
        """
        kwargs.setdefault("timeout", 10)
        for attempt in range(self._MAX_RETRIES + 1):
            try:
                resp = requests.request(method, url, **kwargs)
            except requests.exceptions.RequestException as exc:
                logger.warning("%s: request error — %s", self.name, exc)
                return None
            if resp.status_code in self._RETRY_STATUS and attempt < self._MAX_RETRIES:
                time.sleep(self._BACKOFF_SECS * (attempt + 1))
                continue
            if resp.status_code in self._RETRY_STATUS:
                logger.warning("%s: HTTP %s after retry", self.name, resp.status_code)
            return resp
        return None

    def _norm(self, title="", company="", location="",
              salary_min="", salary_max="", created="",
              url="", description="") -> dict:
        return {
            "title":       title,
            "company":     company,
            "location":    location,
            "salary_min":  salary_min,
            "salary_max":  salary_max,
            "created":     str(created)[:10] if created else "",
            "url":         url,
            "description": str(description).replace("\n", " ")[:300],
            "source":      self.name,
        }
