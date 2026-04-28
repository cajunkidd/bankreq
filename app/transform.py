from collections import Counter, defaultdict
from datetime import date, datetime
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import Border, Font, Side
from openpyxl.workbook import Workbook

SOURCE_SHEET = "Net Sales"
ADDITIONAL_SOURCE_SHEETS = ("Adjustments", "Chargebacks & Chargeback Revers")
OUTPUT_SHEET_BASE = "Formatted"

OUTPUT_HEADERS = [
    "Site Alternate ID",
    "Funded Date",
    "Site Name",
    "Product Code",
    "Processed Transaction Amount",
    None,
]

AMOUNT_FORMAT = "#,###.00;\\-#,###.00"

COLUMN_WIDTHS = {
    "A": 13.0,
    "B": 12.3,
    "C": 22.7,
    "D": 14.0,
    "E": 17.4,
    "F": 14.6,
}

THIN = Side(style="thin")
THICK = Side(style="thick")
NONE_SIDE = Side(style=None)


def _product_sort_key(pc: str) -> str:
    return (pc or "").lower()


REQUIRED_COLUMNS = (
    "Site Alternate ID",
    "Site Name",
    "Funded Date",
    "Product Code",
    "Processed Transaction Amount",
)


def _ingest_sheet(ws, aggregated, site_meta, date_counts) -> bool:
    """Read rows from a single sheet into the shared aggregates. Returns
    True if the sheet had a recognisable header, False otherwise (which
    we treat as 'empty/blank' and skip silently)."""
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not header_row:
        return False
    headers = list(header_row)
    if not set(REQUIRED_COLUMNS).issubset(set(headers)):
        return False
    idx = {name: headers.index(name) for name in REQUIRED_COLUMNS}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        alt_id = row[idx["Site Alternate ID"]]
        sname = row[idx["Site Name"]]
        fdate = row[idx["Funded Date"]]
        pc = row[idx["Product Code"]]
        amt = row[idx["Processed Transaction Amount"]]
        if alt_id is None or pc is None or amt is None:
            continue
        aggregated[(alt_id, pc)] += float(amt)
        site_meta.setdefault(alt_id, (sname, fdate))
        if fdate is not None:
            date_counts[fdate] += 1
    return True


def _read_source_rows(wb: Workbook):
    if SOURCE_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Workbook is missing the required '{SOURCE_SHEET}' sheet. "
            f"Found sheets: {wb.sheetnames}"
        )

    aggregated: dict[tuple, float] = defaultdict(float)
    site_meta: dict[str, tuple[str, str]] = {}
    date_counts: Counter = Counter()

    primary_ok = _ingest_sheet(
        wb[SOURCE_SHEET], aggregated, site_meta, date_counts
    )
    if not primary_ok:
        raise ValueError(
            f"'{SOURCE_SHEET}' sheet is missing required columns: "
            f"{sorted(REQUIRED_COLUMNS)}"
        )

    for sheet_name in ADDITIONAL_SOURCE_SHEETS:
        if sheet_name in wb.sheetnames:
            _ingest_sheet(
                wb[sheet_name], aggregated, site_meta, date_counts
            )

    funded_date = date_counts.most_common(1)[0][0] if date_counts else None
    return aggregated, site_meta, funded_date


def _next_sheet_name(wb: Workbook) -> str:
    if OUTPUT_SHEET_BASE not in wb.sheetnames:
        return OUTPUT_SHEET_BASE
    n = 2
    while f"{OUTPUT_SHEET_BASE} ({n})" in wb.sheetnames:
        n += 1
    return f"{OUTPUT_SHEET_BASE} ({n})"


def _box_border(top: bool, bottom: bool, left: bool, right: bool) -> Border:
    return Border(
        top=THICK if top else NONE_SIDE,
        bottom=THICK if bottom else NONE_SIDE,
        left=THICK if left else NONE_SIDE,
        right=THICK if right else NONE_SIDE,
    )


def _write_output_sheet(wb: Workbook, aggregated, site_meta) -> str:
    name = _next_sheet_name(wb)
    ws = wb.create_sheet(name, 0)
    wb.active = 0

    bold = Font(bold=True)
    header_border = Border(bottom=THIN)
    for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        if header is not None:
            cell.font = bold
            cell.border = header_border

    for letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[letter].width = width

    sites = sorted(site_meta.keys(), key=lambda s: str(s))
    current_row = 2
    for alt_id in sites:
        sname, fdate = site_meta[alt_id]
        products = sorted(
            [pc for (a, pc) in aggregated.keys() if a == alt_id],
            key=_product_sort_key,
        )
        if not products:
            continue
        site_start_row = current_row
        first_non_amex_row: int | None = None
        for pc in products:
            amt = aggregated[(alt_id, pc)]
            ws.cell(row=current_row, column=1, value=alt_id)
            ws.cell(row=current_row, column=2, value=fdate)
            ws.cell(row=current_row, column=3, value=sname)
            ws.cell(row=current_row, column=4, value=pc)
            amt_cell = ws.cell(row=current_row, column=5, value=amt)
            amt_cell.number_format = AMOUNT_FORMAT
            if pc != "Amex" and first_non_amex_row is None:
                first_non_amex_row = current_row
            current_row += 1
        last_row = current_row - 1

        # Subtotal on the last row of the group: SUM of non-Amex rows.
        # Skip when the site has only one row (no breakdown to subtotal).
        if last_row > site_start_row and first_non_amex_row is not None:
            sub_cell = ws.cell(row=last_row, column=6)
            sub_cell.value = f"=SUM(E{first_non_amex_row}:E{last_row})"
            sub_cell.number_format = AMOUNT_FORMAT

        # Draw the compartment: box around non-Amex rows (the bank-deposit
        # portion). Amex rows sit outside the box because they're funded
        # separately by Amex.
        box_start = first_non_amex_row if first_non_amex_row else site_start_row
        if box_start is None or box_start > last_row:
            continue
        for r in range(box_start, last_row + 1):
            top = r == box_start
            bottom = r == last_row
            for c in range(1, 7):
                left = c == 1
                right = c == 6
                ws.cell(row=r, column=c).border = _box_border(top, bottom, left, right)

    return name


def _normalize_date(raw) -> date:
    """Return a date object from the funded-date cell value, falling back to
    today if it can't be parsed."""
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    if isinstance(raw, str):
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(raw, fmt).date()
            except ValueError:
                continue
    return date.today()


def reformat_workbook(file_bytes: bytes) -> tuple[bytes, str, date]:
    """Take a raw merchant-services workbook, append a 'Formatted' sheet
    (auto-suffixed if one already exists), and return
    (workbook_bytes, sheet_name, funded_date)."""
    wb = load_workbook(BytesIO(file_bytes))
    aggregated, site_meta, raw_date = _read_source_rows(wb)
    sheet_name = _write_output_sheet(wb, aggregated, site_meta)
    out = BytesIO()
    wb.save(out)
    return out.getvalue(), sheet_name, _normalize_date(raw_date)
