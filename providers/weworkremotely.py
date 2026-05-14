import requests
import xml.etree.ElementTree as ET
import re
from email.utils import parsedate
from .base import BaseProvider


def _parse_rss_date(s):
    try:
        t = parsedate(s)
        if t:
            return f"{t[0]}-{t[1]:02d}-{t[2]:02d}"
    except Exception:
        pass
    return ""

def _strip_html(text):
    return re.sub(r"<[^>]+>", " ", text or "").strip()


_FEEDS = [
    "https://weworkremotely.com/remote-jobs.rss",
    "https://weworkremotely.com/categories/remote-full-stack-programming-jobs.rss",
    "https://weworkremotely.com/categories/remote-front-end-programming-jobs.rss",
    "https://weworkremotely.com/categories/remote-back-end-programming-jobs.rss",
    "https://weworkremotely.com/categories/remote-devops-sysadmin-jobs.rss",
    "https://weworkremotely.com/categories/remote-management-and-finance-jobs.rss",
    "https://weworkremotely.com/categories/remote-product-jobs.rss",
    "https://weworkremotely.com/categories/remote-sales-and-marketing-jobs.rss",
    "https://weworkremotely.com/categories/remote-design-jobs.rss",
    "https://weworkremotely.com/categories/remote-customer-support-jobs.rss",
    "https://weworkremotely.com/categories/remote-other-jobs.rss",
]


class WeWorkRemotelyProvider(BaseProvider):
    name = "WWR"

    def search(self, what, where, country="us", results=20, **kwargs):
        words     = what.lower().split() if what.strip() else []
        seen_guids = set()
        collected  = []
        try:
            for feed_url in _FEEDS:
                resp = requests.get(
                    feed_url,
                    headers={"User-Agent": "JobSearchTool/1.0"},
                    timeout=15,
                )
                if resp.status_code != 200:
                    continue
                root = ET.fromstring(resp.content)
                for item in root.findall(".//item"):
                    guid = item.findtext("guid", "")
                    if guid in seen_guids:
                        continue
                    seen_guids.add(guid)
                    if words:
                        title    = _strip_html(item.findtext("title", ""))
                        desc     = _strip_html(item.findtext("description", ""))
                        haystack = (title + " " + desc).lower()
                        if not all(w in haystack for w in words):
                            continue
                    collected.append(self._normalize(item))
                if len(collected) >= results:
                    break
        except Exception as e:
            print(f"[WWR] Error: {e}")
        return collected[:results]

    def _normalize(self, item):
        title   = _strip_html(item.findtext("title", ""))
        company = ""
        if ": " in title:
            company, title = title.split(": ", 1)
        url = ""
        for child in item:
            if child.tag == "link" and child.text:
                url = child.text.strip()
            elif child.tag == "guid" and not url:
                url = child.text or ""
        return self._norm(
            title       = title,
            company     = company,
            location    = item.findtext("region", "") or "Remote",
            url         = url,
            created     = _parse_rss_date(item.findtext("pubDate", "")),
            description = _strip_html(item.findtext("description", "")),
        )
