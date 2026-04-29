"""Commodity lumber spot price (Lumber Futures, ticker LBR=F).

Pulled from Yahoo Finance's public chart endpoint — free, no API key.
Data is delayed ~15 min, which is fine for an AP-tools dashboard.
"""
import time
from typing import Optional

import httpx

SYMBOL = "LBR=F"  # CME Lumber Futures (replaced delisted Random Length contract)
URL = (
    "https://query1.finance.yahoo.com/v8/finance/chart/"
    f"{SYMBOL}?interval=1d&range=1mo"
)
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0 Safari/537.36"
)
TIMEOUT = httpx.Timeout(5.0, connect=3.0)
CACHE_TTL_SECONDS = 10 * 60  # 10 minutes — futures move slowly enough

_cache: dict[str, tuple[float, dict]] = {}


def _fetch() -> Optional[dict]:
    headers = {"User-Agent": UA, "Accept": "application/json"}
    try:
        with httpx.Client(timeout=TIMEOUT, headers=headers, follow_redirects=True) as c:
            r = c.get(URL)
            r.raise_for_status()
            payload = r.json()
    except Exception:
        return None

    try:
        result = payload["chart"]["result"][0]
        meta = result["meta"]
        price = meta.get("regularMarketPrice")
        prev = meta.get("chartPreviousClose") or meta.get("previousClose")
        currency = meta.get("currency", "USD")
        closes = result["indicators"]["quote"][0].get("close") or []
        # Drop trailing nulls (markets closed today, holidays, etc.)
        sparkline = [c for c in closes if c is not None][-10:]
    except (KeyError, IndexError, TypeError):
        return None

    if price is None or prev is None:
        return None

    change = price - prev
    pct = (change / prev) * 100 if prev else 0.0
    return {
        "symbol": SYMBOL,
        "name": "Lumber Futures",
        "price": round(price, 2),
        "previous_close": round(prev, 2),
        "change": round(change, 2),
        "percent": round(pct, 2),
        "currency": currency,
        "sparkline": [round(c, 2) for c in sparkline],
    }


def get_quote() -> Optional[dict]:
    cached = _cache.get("quote")
    if cached and time.time() < cached[0]:
        return cached[1]
    q = _fetch()
    if q is None:
        return None
    _cache["quote"] = (time.time() + CACHE_TTL_SECONDS, q)
    return q
