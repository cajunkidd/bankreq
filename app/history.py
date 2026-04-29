"""Local persistence of past uploads, used to flag anomalies on
subsequent uploads.

Storage: SQLite at %APPDATA%/StineBankReq/history.db on Windows, or
~/.bankreq/history.db elsewhere. The directory is created on first use,
so users don't have to set anything up.
"""
import os
import sqlite3
import statistics
import sys
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable, Optional


def data_dir() -> Path:
    if sys.platform == "win32" or getattr(sys, "frozen", False):
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / "StineBankReq"
    return Path.home() / ".bankreq"


def db_path() -> Path:
    d = data_dir()
    d.mkdir(parents=True, exist_ok=True)
    return d / "history.db"


@contextmanager
def _conn():
    con = sqlite3.connect(db_path())
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    finally:
        con.close()


def init_db() -> None:
    with _conn() as con:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS uploads (
                funded_date TEXT NOT NULL,
                site_alternate_id TEXT NOT NULL,
                site_name TEXT,
                product_code TEXT NOT NULL,
                source_sheet TEXT NOT NULL,
                amount REAL NOT NULL,
                ingested_at TEXT DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (funded_date, site_alternate_id, product_code, source_sheet)
            )
            """
        )
        con.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_history_lookup
                ON uploads (site_alternate_id, product_code, source_sheet, funded_date)
            """
        )


def record_upload(
    rows: Iterable[tuple[str, Optional[str], str, str, float]],
    funded_date: date,
) -> None:
    """rows is an iterable of
    (site_alternate_id, site_name, product_code, source_sheet, amount).
    Re-uploads of the same day overwrite the prior value (ON CONFLICT
    REPLACE), so users can re-process a file without doubling up."""
    init_db()
    with _conn() as con:
        for sid, sname, pc, source, amt in rows:
            con.execute(
                """
                INSERT INTO uploads
                  (funded_date, site_alternate_id, site_name, product_code, source_sheet, amount)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(funded_date, site_alternate_id, product_code, source_sheet)
                DO UPDATE SET amount = excluded.amount,
                              site_name = excluded.site_name,
                              ingested_at = CURRENT_TIMESTAMP
                """,
                (funded_date.isoformat(), str(sid), sname, pc, source, float(amt)),
            )


@dataclass(frozen=True)
class Baseline:
    mean: float
    stdev: float
    n: int

    def lower(self, k: float = 2.0) -> float:
        return self.mean - k * self.stdev

    def upper(self, k: float = 2.0) -> float:
        return self.mean + k * self.stdev


def baseline(
    site_alternate_id: str,
    product_code: str,
    source_sheet: str,
    asof: date,
    lookback_days: int = 60,
    min_points: int = 5,
) -> Optional[Baseline]:
    """Return a Baseline computed from rows in the trailing `lookback_days`
    days strictly before `asof`. Returns None if there aren't yet enough
    data points to be statistically meaningful."""
    init_db()
    start = (asof - timedelta(days=lookback_days)).isoformat()
    end = asof.isoformat()
    with _conn() as con:
        cur = con.execute(
            """
            SELECT amount FROM uploads
            WHERE site_alternate_id = ? AND product_code = ? AND source_sheet = ?
              AND funded_date >= ? AND funded_date < ?
            """,
            (str(site_alternate_id), product_code, source_sheet, start, end),
        )
        amounts = [r["amount"] for r in cur.fetchall()]
    if len(amounts) < min_points:
        return None
    mean = statistics.mean(amounts)
    stdev = statistics.stdev(amounts) if len(amounts) > 1 else 0.0
    return Baseline(mean=mean, stdev=stdev, n=len(amounts))
