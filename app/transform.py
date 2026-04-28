from collections import defaultdict
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

SOURCE_SHEET = "Net Sales"
OUTPUT_SHEET = "Net Sales (2)"

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

# Stable order: alphabetical, which matches the team's existing convention
# (Amex, DebitCard, Discover, Mastercard, Miscellaneous, Visa).
def _product_sort_key(pc: str) -> str:
    return (pc or "").lower()


def _read_source_rows(wb: Workbook):
    if SOURCE_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Workbook is missing the required '{SOURCE_SHEET}' sheet. "
            f"Found sheets: {wb.sheetnames}"
        )
    ws = wb[SOURCE_SHEET]
    headers = [c.value for c in ws[1]]
    required = {
        "Site Alternate ID",
        "Site Name",
        "Funded Date",
        "Product Code",
        "Processed Transaction Amount",
    }
    missing = required - set(headers)
    if missing:
        raise ValueError(
            f"'{SOURCE_SHEET}' sheet is missing required columns: {sorted(missing)}"
        )
    idx = {name: headers.index(name) for name in required}

    aggregated: dict[tuple, float] = defaultdict(float)
    site_meta: dict[str, tuple[str, str]] = {}
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

    return aggregated, site_meta


def _write_output_sheet(wb: Workbook, aggregated, site_meta) -> None:
    if OUTPUT_SHEET in wb.sheetnames:
        del wb[OUTPUT_SHEET]
    ws = wb.create_sheet(OUTPUT_SHEET)

    bold = Font(bold=True)
    for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        if header is not None:
            cell.font = bold

    for letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[letter].width = width

    sites = sorted(site_meta.keys(), key=lambda s: (str(s)))
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
        # Skip when the site has only one row (matches the sample).
        if last_row > site_start_row and first_non_amex_row is not None:
            sub_cell = ws.cell(row=last_row, column=6)
            sub_cell.value = f"=SUM(E{first_non_amex_row}:E{last_row})"
            sub_cell.number_format = AMOUNT_FORMAT


def reformat_workbook(file_bytes: bytes) -> bytes:
    """Take a raw merchant-services workbook, add the formatted Net Sales (2)
    sheet, and return the resulting workbook as bytes."""
    wb = load_workbook(BytesIO(file_bytes))
    aggregated, site_meta = _read_source_rows(wb)
    _write_output_sheet(wb, aggregated, site_meta)
    out = BytesIO()
    wb.save(out)
    return out.getvalue()
