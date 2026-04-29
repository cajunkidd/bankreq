"""US Treasury yields tracker (free, via Yahoo Finance)."""
import concurrent.futures
import time
from typing import Optional

import httpx

# (label, yahoo symbol). ^IRX is the 13-week T-bill (short-term proxy
# for Fed Funds direction); ^TNX is the 10-year Treasury yield (the
# benchmark mortgage rates track); ^TYX is the 30-year Treasury yield.
RATES: list[tuple[str, str]] = [
    ("3-Month",  "^IRX"),
    ("10-Year",  "^TNX"),
    ("30-Year",  "^TYX"),
]

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0 Safari/537.36"
)
TIMEOUT = httpx.Timeout(5.0, connect=3.0)
CACHE_TTL_SECONDS = 10 * 60

_cache: dict[str, tuple[float, list]] = {}


def _fetch_one(label: str, symbol: str) -> Optional[dict]:
    url = (
        "https://query1.finance.yahoo.com/v8/finance/chart/"
        f"{symbol}?interval=1d&range=5d"
    )
    headers = {"User-Agent": UA, "Accept": "application/json"}
    try:
        with httpx.Client(timeout=TIMEOUT, headers=headers, follow_redirects=True) as c:
            r = c.get(url)
            r.raise_for_status()
            payload = r.json()
        meta = payload["chart"]["result"][0]["meta"]
        price = meta.get("regularMarketPrice")
        prev = meta.get("chartPreviousClose") or meta.get("previousClose")
        if price is None or prev is None:
            return None
    except Exception:
        return None
    change = price - prev
    return {
        "label": label,
        "symbol": symbol,
        "yield": round(price, 2),
        "previous_close": round(prev, 2),
        "change_bps": round(change * 100, 1),  # 1bp = 0.01%
    }


def get_rates() -> list[dict]:
    cached = _cache.get("rates")
    if cached and time.time() < cached[0]:
        return cached[1]
    results: list[dict] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(RATES)) as pool:
        futures = {pool.submit(_fetch_one, label, sym): label for label, sym in RATES}
        # Preserve order from RATES
        by_label: dict[str, dict] = {}
        for fut in concurrent.futures.as_completed(futures, timeout=8):
            try:
                row = fut.result()
                if row:
                    by_label[row["label"]] = row
            except Exception:
                continue
        for label, _sym in RATES:
            if label in by_label:
                results.append(by_label[label])
    _cache["rates"] = (time.time() + CACHE_TTL_SECONDS, results)
    return results
