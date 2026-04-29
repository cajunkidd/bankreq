"""Sulphur, LA weather via the National Weather Service public API.

Free, no API key, no quota. weather.gov requires a User-Agent header.
"""
import time
from typing import Optional

import httpx

# Sulphur, LA — National Weather Service uses lat/lon -> gridpoint -> forecast.
SULPHUR_LA_LAT = 30.2367
SULPHUR_LA_LON = -93.3771

UA = "StineBankReq/1.0 (Stine IT internal tool)"
TIMEOUT = httpx.Timeout(5.0, connect=3.0)
CACHE_TTL_SECONDS = 15 * 60  # 15 min — NWS forecasts only update hourly anyway

_cache: dict[str, tuple[float, dict]] = {}


def _cached(key: str) -> Optional[dict]:
    hit = _cache.get(key)
    if hit is None:
        return None
    expires_at, value = hit
    if time.time() > expires_at:
        _cache.pop(key, None)
        return None
    return value


def _store(key: str, value: dict) -> None:
    _cache[key] = (time.time() + CACHE_TTL_SECONDS, value)


def _fetch_forecast() -> dict:
    headers = {"User-Agent": UA, "Accept": "application/geo+json"}
    with httpx.Client(timeout=TIMEOUT, headers=headers) as client:
        # Step 1: resolve lat/lon to a forecast endpoint URL.
        points_url = f"https://api.weather.gov/points/{SULPHUR_LA_LAT},{SULPHUR_LA_LON}"
        r = client.get(points_url)
        r.raise_for_status()
        forecast_url = r.json()["properties"]["forecast"]
        # Step 2: pull the forecast.
        f = client.get(forecast_url)
        f.raise_for_status()
        return f.json()


def get_forecast() -> Optional[dict]:
    """Return a small dict suitable for the landing-page widget, or None
    if the API is unreachable. Result shape:
        {
          "location": "Sulphur, LA",
          "periods": [{"name", "short", "temp", "icon"}, ...]   # next 4
        }
    """
    cached = _cached("forecast")
    if cached is not None:
        return cached
    try:
        raw = _fetch_forecast()
    except Exception:
        return None

    periods = raw.get("properties", {}).get("periods", [])[:4]
    out = {
        "location": "Sulphur, LA",
        "periods": [
            {
                "name": p.get("name"),
                "short": p.get("shortForecast"),
                "temp": p.get("temperature"),
                "unit": p.get("temperatureUnit", "F"),
                "icon": p.get("icon"),
                "wind": f"{p.get('windSpeed','')} {p.get('windDirection','')}".strip(),
            }
            for p in periods
        ],
    }
    _store("forecast", out)
    return out
