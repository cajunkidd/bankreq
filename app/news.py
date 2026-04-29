"""Aggregated LBM (lumber & building materials) industry news, parsed
from the public RSS / Atom feeds of trade publications.

All sources are free and don't require an API key. Each feed is fetched
concurrently with a short timeout; any that fails (404, network error,
malformed XML) is skipped silently so the widget always renders
something. Parsing is done with the stdlib's xml.etree to avoid pulling
in a third-party RSS dependency.
"""
import concurrent.futures
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from typing import Optional

import httpx

# Trade pubs covering the LBM / building-materials industry. If any of
# these break (URL changes, server error), the rest still load.
FEEDS: list[tuple[str, str]] = [
    ("LBM Journal",       "https://lbmjournal.com/feed/"),
    ("HBS Dealer",        "https://www.hbsdealer.com/rss.xml"),
    ("ProSales",          "https://www.prosalesmagazine.com/rss/"),
    ("Construction Dive", "https://www.constructiondive.com/feeds/news/"),
    ("Builder Online",    "https://www.builderonline.com/rss"),
    ("Remodeling",        "https://www.remodeling.hw.net/rss"),
]

UA = "StineBankReq/1.0 (Stine IT internal tool)"
TIMEOUT = httpx.Timeout(5.0, connect=3.0)
CACHE_TTL_SECONDS = 15 * 60
MAX_PER_SOURCE = 3
MAX_TOTAL = 8

ATOM_NS = "{http://www.w3.org/2005/Atom}"

_cache: dict[str, tuple[float, list]] = {}


def _parse_date(s: Optional[str]) -> float:
    """Parse RFC 822 (RSS pubDate) or ISO-8601 (Atom updated/published)
    into epoch seconds. Returns 0 if unparseable."""
    if not s:
        return 0.0
    try:
        dt = parsedate_to_datetime(s)
        if dt is not None:
            return dt.timestamp()
    except Exception:
        pass
    try:
        # Atom uses ISO-8601, sometimes with trailing Z.
        s2 = s.replace("Z", "+00:00")
        return datetime.fromisoformat(s2).timestamp()
    except Exception:
        return 0.0


def _parse_feed(content: bytes) -> list[dict]:
    try:
        root = ET.fromstring(content)
    except ET.ParseError:
        return []

    items: list[dict] = []

    # RSS 2.0: <rss><channel><item><title/><link/><pubDate/></item>
    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        pub = item.findtext("pubDate") or item.findtext("{http://purl.org/dc/elements/1.1/}date")
        if title and link:
            items.append({"title": title, "url": link, "ts": _parse_date(pub)})

    # Atom: <feed><entry><title/><link href=".."/><updated/></entry>
    if not items:
        for entry in root.iter(f"{ATOM_NS}entry"):
            title_el = entry.find(f"{ATOM_NS}title")
            title = (title_el.text or "").strip() if title_el is not None else ""
            link = ""
            for link_el in entry.findall(f"{ATOM_NS}link"):
                rel = link_el.attrib.get("rel", "alternate")
                if rel == "alternate":
                    link = link_el.attrib.get("href", "")
                    break
            if not link:
                first = entry.find(f"{ATOM_NS}link")
                if first is not None:
                    link = first.attrib.get("href", "")
            ts = _parse_date(
                entry.findtext(f"{ATOM_NS}updated")
                or entry.findtext(f"{ATOM_NS}published")
            )
            if title and link:
                items.append({"title": title, "url": link, "ts": ts})

    return items


def _fetch_one(name: str, url: str) -> list[dict]:
    headers = {"User-Agent": UA, "Accept": "application/rss+xml, application/atom+xml, application/xml;q=0.9, */*;q=0.8"}
    try:
        with httpx.Client(timeout=TIMEOUT, headers=headers, follow_redirects=True) as c:
            r = c.get(url)
            r.raise_for_status()
            entries = _parse_feed(r.content)
    except Exception:
        return []
    out: list[dict] = []
    for e in entries[:MAX_PER_SOURCE]:
        out.append({
            "title": e["title"],
            "url": e["url"],
            "source": name,
            "ts": e["ts"],
            "published_iso": (
                datetime.fromtimestamp(e["ts"], tz=timezone.utc).isoformat()
                if e["ts"] else None
            ),
        })
    return out


def get_headlines() -> list[dict]:
    """Return the latest headlines across all configured feeds, sorted
    by publish date desc, capped at MAX_TOTAL. Empty list on total
    failure (so the widget can decide to hide itself)."""
    cached = _cache.get("headlines")
    if cached and time.time() < cached[0]:
        return cached[1]

    all_items: list[dict] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(FEEDS)) as pool:
        futures = [pool.submit(_fetch_one, name, url) for name, url in FEEDS]
        for f in concurrent.futures.as_completed(futures, timeout=8):
            try:
                all_items.extend(f.result() or [])
            except Exception:
                continue

    all_items.sort(key=lambda x: x["ts"], reverse=True)
    top = all_items[:MAX_TOTAL]
    _cache["headlines"] = (time.time() + CACHE_TTL_SECONDS, top)
    return top
