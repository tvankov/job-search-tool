from abc import ABC, abstractmethod


class AuthError(Exception):
    """Raised when a provider rejects the API key (HTTP 401 / 403)."""


class BaseProvider(ABC):
    name: str = ""

    @abstractmethod
    def search(self, what: str, where: str, country: str,
               results: int, **kwargs) -> list:
        """Return list of normalized job dicts."""

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
