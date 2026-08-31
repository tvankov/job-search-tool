from .base             import AuthError
from .adzuna           import AdzunaProvider
from .reed             import ReedProvider
from .findwork         import FindworkProvider
from .arbeitnow        import ArbeitnowProvider
from .remoteok         import RemoteOKProvider
from .jooble           import JoobleProvider
from .themuse          import TheMuseProvider
from .bundesagentur    import BundesagenturProvider
from .headhunter       import HeadHunterProvider
from .weworkremotely   import WeWorkRemotelyProvider
from .remotive         import RemotiveProvider
from .himalayas        import HimalayasProvider

# NOTE: CareerJet and Recruit (JP) are intentionally NOT wired in — their APIs
# did not work reliably in testing, so they are excluded from the active
# search path. The modules are kept for reference only.

__all__ = [
    "AuthError",
    "AdzunaProvider", "ReedProvider", "FindworkProvider",
    "ArbeitnowProvider", "RemoteOKProvider", "JoobleProvider",
    "TheMuseProvider", "BundesagenturProvider",
    "HeadHunterProvider", "WeWorkRemotelyProvider", "RemotiveProvider",
    "HimalayasProvider",
    "build_provider_tasks", "dedup_jobs",
]


def build_provider_tasks(*, use, creds, what, where, country, results,
                         sort_by="date", salary_min=None, salary_max=None,
                         full_time=False, permanent=False):
    """Return ``[(name, callable), ...]`` for every enabled provider that has
    the credentials it requires.

    This is the single source of truth for wiring config -> provider calls,
    shared by the GUI (`job_search_4._search`) and the headless auto-runner
    (`job_search_auto.fetch_jobs`). The list order defines result priority
    once duplicate URLs are merged (see :func:`dedup_jobs`).

    ``use``   maps a provider key (e.g. "adzuna") to a bool (enabled).
    ``creds`` holds API keys/ids: adzuna_id, adzuna_key, reed_key,
              findwork_key, jooble_key, hh_token, themuse_key.
    """
    candidates = [
        # ── Providers that require a key/id (skipped if missing) ──────────────
        ("Adzuna", lambda: AdzunaProvider(
            creds.get("adzuna_id", ""), creds.get("adzuna_key", ""),
        ).search(what, where, country=country, results=results, sort_by=sort_by,
                 salary_min=salary_min, salary_max=salary_max,
                 full_time=full_time, permanent=permanent),
            use.get("adzuna", False) and bool(creds.get("adzuna_id") and creds.get("adzuna_key"))),
        ("Reed", lambda: ReedProvider(
            creds.get("reed_key", "")).search(what, where, results=results),
            use.get("reed", False) and bool(creds.get("reed_key"))),
        ("Findwork", lambda: FindworkProvider(
            creds.get("findwork_key", "")).search(what, where, results=results),
            use.get("findwork", False) and bool(creds.get("findwork_key"))),
        ("Jooble", lambda: JoobleProvider(
            creds.get("jooble_key", "")).search(what, where, results=results),
            use.get("jooble", False) and bool(creds.get("jooble_key"))),
        ("HeadHunter", lambda: HeadHunterProvider(
            creds.get("hh_token", "")).search(what, where, results=results),
            use.get("headhunter", False) and bool(creds.get("hh_token"))),
        # ── Keyless providers (enabled by default) ───────────────────────────
        ("Arbeitnow", lambda: ArbeitnowProvider().search(
            what, where, results=results),
            use.get("arbeitnow", True)),
        ("Bundesagentur", lambda: BundesagenturProvider().search(
            what, where, results=results),
            use.get("bundesag", True)),
        ("RemoteOK", lambda: RemoteOKProvider().search(
            what, where, results=results),
            use.get("remoteok", True)),
        ("The Muse", lambda: TheMuseProvider(
            creds.get("themuse_key", "")).search(what, where, results=results),
            use.get("themuse", True)),
        ("WeWorkRemotely", lambda: WeWorkRemotelyProvider().search(
            what, where, results=results),
            use.get("wwr", True)),
        ("Remotive", lambda: RemotiveProvider().search(
            what, where, results=results),
            use.get("remotive", True)),
        ("Himalayas", lambda: HimalayasProvider().search(
            what, where, results=results),
            use.get("himalayas", True)),
    ]
    return [(name, fn) for name, fn, enabled in candidates if enabled]


def dedup_jobs(job_lists, cap=None):
    """Merge job lists (given in priority order) into one, dropping duplicate
    URLs. The first occurrence of a URL wins. Stops early once ``cap`` items
    are collected (``None`` = no limit)."""
    seen = set()
    out  = []
    for jobs in job_lists:
        for j in jobs:
            if cap is not None and len(out) >= cap:
                return out
            url = j.get("url", "")
            if url and url in seen:
                continue
            if url:
                seen.add(url)
            out.append(j)
    return out
