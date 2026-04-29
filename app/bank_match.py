"""Bank-statement matching against the formatted CardConnect output.

This module is scaffolded: the public API and data shapes are defined,
but the parsing of the bank statement is stubbed because we don't yet
have a sample of the AP team's bank statement format. Once a sample
lands, fill in `parse_bank_statement` and `match_deposits` — the rest
of the pipeline (route, response shape, UI surface) is already there.

Expected matching logic (to be confirmed once we have a sample):
  - For each site in the formatted output, take the per-site subtotal
    (the SUM in column F: bank-deposit total excluding Amex).
  - For each deposit line on the bank statement, find a site whose
    subtotal matches within $0.01 on the funded date (or +/- 2 days
    to allow for processor lag).
  - Surface three buckets: matched, unmatched bank lines, unmatched
    site totals.
"""
from dataclasses import dataclass, field
from datetime import date
from typing import Optional


@dataclass(frozen=True)
class BankLine:
    posted_date: date
    amount: float
    description: str
    reference: Optional[str] = None


@dataclass(frozen=True)
class SiteTotal:
    funded_date: date
    site_alternate_id: str
    site_name: Optional[str]
    amount: float  # excludes Amex (matches the bank deposit)


@dataclass
class MatchResult:
    matched: list[tuple[SiteTotal, BankLine]] = field(default_factory=list)
    unmatched_bank: list[BankLine] = field(default_factory=list)
    unmatched_sites: list[SiteTotal] = field(default_factory=list)


def parse_bank_statement(file_bytes: bytes) -> list[BankLine]:
    """Parse a bank-statement file into structured BankLine rows.

    TODO: implement once we have a sample of the AP team's bank
    statement. Supported formats will likely be Excel (.xlsx) and CSV.
    """
    raise NotImplementedError(
        "Bank-statement parsing is not yet implemented — please share a "
        "sample bank statement so we can encode the column layout."
    )


def site_totals_from_workbook(file_bytes: bytes) -> list[SiteTotal]:
    """Pull the per-site bank-deposit totals out of a Formatted workbook.

    TODO: walk the Formatted sheet, find the SUM cell on the last row of
    each site box, and return one SiteTotal per box.
    """
    raise NotImplementedError(
        "Site-total extraction is not yet implemented."
    )


def match_deposits(
    site_totals: list[SiteTotal],
    bank_lines: list[BankLine],
    *,
    amount_tolerance: float = 0.01,
    date_window_days: int = 2,
) -> MatchResult:
    """Match bank lines to site totals.

    TODO: implement greedy match by (date +/- window, amount within
    tolerance). For now, return everything as unmatched so callers can
    smoke-test the wiring."""
    return MatchResult(
        matched=[],
        unmatched_bank=list(bank_lines),
        unmatched_sites=list(site_totals),
    )
