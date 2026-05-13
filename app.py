"""Stine BankReq Reformatter — drag-and-drop the daily raw transaction file and
get back the reformatted BankReq workbook with a new Formatted sheet."""
from __future__ import annotations

import base64
import csv
import hashlib
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, Side
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).parent
LOGO_FILE = ROOT / "Stinelogo_white_rec.svg"

st.set_page_config(page_title="Stine CardConnect Reformatter", layout="wide")

st.markdown(
    """
    <style>
    .stDeployButton {display: none !important;}
    [data-testid="stDeployButton"] {display: none !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

OUT_COLS = [
    "Site Alternate ID",
    "Funded Date",
    "Site Name",
    "Product Code",
    "Processed Transaction Amount",
]

PRODUCT_ORDER = ["Amex", "DebitCard", "Discover", "Mastercard", "Visa"]

NET_SALES_SHEET = "Net Sales"
ADJUSTMENTS_SHEET = "Adjustments"
CHARGEBACKS_SHEET = "Chargebacks & Chargeback Revers"
RED = "FFFF0000"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

THICK = Side(style="thick")
THIN = Side(style="thin")
BOLD = Font(bold=True)
RED_FONT = Font(color=RED)


def _product_key(p: object) -> tuple[int, int, str]:
    s = str(p)
    if s in PRODUCT_ORDER:
        return (0, PRODUCT_ORDER.index(s), s)
    return (1, 0, s)


def _funded_date_str(value: object) -> str:
    if isinstance(value, str):
        return value
    if pd.isna(value):
        return ""
    try:
        return pd.to_datetime(value).strftime("%m/%d/%Y")
    except Exception:
        return str(value)


def _site_key(value: object) -> tuple[int, float, str]:
    s = str(value)
    try:
        return (0, float(s), s)
    except (TypeError, ValueError):
        return (1, 0.0, s)


def _group_section(df: pd.DataFrame) -> pd.DataFrame:
    needed = {
        "Site Alternate ID",
        "Funded Date",
        "Site Name",
        "Product Code",
        "Processed Transaction Amount",
    }
    if not needed.issubset(df.columns):
        return pd.DataFrame(columns=OUT_COLS)

    work = df[list(needed)].copy()
    work["Funded Date"] = work["Funded Date"].map(_funded_date_str)
    grouped = (
        work.groupby(
            ["Site Alternate ID", "Funded Date", "Site Name", "Product Code"],
            dropna=False,
            as_index=False,
        )["Processed Transaction Amount"]
        .sum()
    )
    grouped["__site"] = grouped["Site Alternate ID"].map(_site_key)
    grouped["__prod"] = grouped["Product Code"].map(_product_key)
    grouped = grouped.sort_values(["__site", "__prod"]).drop(columns=["__site", "__prod"])
    return grouped[OUT_COLS].reset_index(drop=True)


def build_branches(sheets: dict[str, pd.DataFrame]) -> list[dict]:
    """Merge every site's Net Sales, Adjustments, and Chargebacks into a single
    branch dict. Adjustments and chargeback rows are tagged red so the writer
    can render them in red text."""
    sources = [
        (NET_SALES_SHEET, "net"),
        (ADJUSTMENTS_SHEET, "adj"),
        (CHARGEBACKS_SHEET, "cb"),
    ]
    by_branch: dict[tuple, dict] = {}
    for sheet_name, role in sources:
        df = sheets.get(sheet_name)
        if df is None or df.empty:
            continue
        block = _group_section(df)
        for _, row in block.iterrows():
            key = (str(row["Site Alternate ID"]), row["Funded Date"], row["Site Name"])
            if key not in by_branch:
                by_branch[key] = {
                    "Site Alternate ID": key[0],
                    "Funded Date": key[1],
                    "Site Name": key[2],
                    "net": [],
                    "adj": [],
                    "cb": [],
                }
            by_branch[key][role].append(
                (row["Product Code"], row["Processed Transaction Amount"])
            )
    return sorted(by_branch.values(), key=lambda b: _site_key(b["Site Alternate ID"]))


def derive_output_filename(branches: list[dict]) -> str:
    for branch in branches:
        raw = branch.get("Funded Date")
        if not raw:
            continue
        try:
            dt = pd.to_datetime(raw)
            return f"BankReq - {dt.strftime('%m-%d-%Y')}.xlsx"
        except Exception:
            break
    return f"BankReq - {datetime.now().strftime('%m-%d-%Y')}.xlsx"


def _apply_border(cell, *, left=None, right=None, top=None, bottom=None) -> None:
    b = cell.border
    cell.border = Border(
        left=left if left is not None else b.left,
        right=right if right is not None else b.right,
        top=top if top is not None else b.top,
        bottom=bottom if bottom is not None else b.bottom,
    )


def _draw_thick_box(ws: Worksheet, top_row: int, bottom_row: int, left_col: int = 1, right_col: int = 6) -> None:
    for c in range(left_col, right_col + 1):
        _apply_border(ws.cell(row=top_row, column=c), top=THICK)
        _apply_border(ws.cell(row=bottom_row, column=c), bottom=THICK)
    for r in range(top_row, bottom_row + 1):
        _apply_border(ws.cell(row=r, column=left_col), left=THICK)
        _apply_border(ws.cell(row=r, column=right_col), right=THICK)


def _write_data_row(
    ws: Worksheet,
    row: int,
    branch: dict,
    product: str,
    amount: object,
    *,
    red: bool = False,
) -> None:
    site_id = branch["Site Alternate ID"]
    cells = [
        ws.cell(row=row, column=1, value=str(site_id) if site_id is not None else None),
        ws.cell(row=row, column=2, value=branch["Funded Date"]),
        ws.cell(row=row, column=3, value=branch["Site Name"]),
        ws.cell(row=row, column=4, value=product),
        ws.cell(
            row=row,
            column=5,
            value=float(amount) if pd.notna(amount) else None,
        ),
    ]
    if red:
        for c in cells:
            c.font = RED_FONT


def _write_branch(ws: Worksheet, start_row: int, branch: dict) -> int:
    net_rows = branch["net"]
    amex = [(p, a) for p, a in net_rows if str(p) == "Amex"]
    others = [(p, a) for p, a in net_rows if str(p) != "Amex"]
    others.sort(key=lambda x: _product_key(x[0]))

    cur = start_row
    for product, amount in amex:
        _write_data_row(ws, cur, branch, product, amount)
        cur += 1

    box_top = cur
    for product, amount in others:
        _write_data_row(ws, cur, branch, product, amount)
        cur += 1
    for product, amount in branch["adj"]:
        _write_data_row(ws, cur, branch, product, amount, red=True)
        cur += 1
    for product, amount in branch["cb"]:
        _write_data_row(ws, cur, branch, product, amount, red=True)
        cur += 1

    if cur > box_top:
        box_bottom = cur - 1
        _draw_thick_box(ws, box_top, box_bottom)
        ws.cell(
            row=box_bottom,
            column=6,
            value=f"=SUM(E{box_top}:E{box_bottom})",
        )

    return cur


def write_workbook(raw_bytes: bytes, branches: list[dict]) -> bytes:
    wb = load_workbook(io.BytesIO(raw_bytes))
    if "Formatted" in wb.sheetnames:
        del wb["Formatted"]
    ws = wb.create_sheet("Formatted", 0)

    for col_idx, name in enumerate(OUT_COLS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = BOLD
        _apply_border(cell, bottom=THIN)

    cur = 2
    for i, branch in enumerate(branches):
        if i > 0:
            cur += 1  # blank row between branches for readability
        cur = _write_branch(ws, cur, branch)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def trigger_browser_download(data: bytes, filename: str, token: str) -> None:
    if st.session_state.get("_last_auto_download") == token:
        return
    st.session_state["_last_auto_download"] = token
    b64 = base64.b64encode(data).decode("ascii")
    components.html(
        f"""
        <script>
          const a = document.createElement('a');
          a.href = "data:{XLSX_MIME};base64,{b64}";
          a.download = {filename!r};
          document.body.appendChild(a);
          a.click();
          a.remove();
        </script>
        """,
        height=0,
    )


# ---------------------------------------------------------------------------
# Wells Fargo Funding reformatter
# ---------------------------------------------------------------------------

WF_OUT_COLS = [
    "Date",
    "Tran Dt",
    "Bank",
    "Acct",
    "Name",
    "Ticket",
    "Gross-Chg",
    "Discount",
    "Net Charge",
]
WF_RAW_COLS = [
    "Date",
    "Tran Dt",
    "Bank",
    "Merchant Num",
    "Key",
    "Acct",
    "Plan",
    "Des",
    "Name",
    "Ticket",
    "Gross-Chg",
    "Discount",
    "Net Charge",
]
WF_MONEY_FMT = '"$"#,##0.00_);[Red]("$"#,##0.00)'
WF_DATE_FMT = "mm-dd-yy"
WF_COL_WIDTHS = {"A": 11.43, "B": 9.71, "E": 18.86, "F": 10.0, "G": 9.86, "I": 10.86}
WF_SUM_FONT = Font(bold=True, size=14)
WF_SUM_ROW_HEIGHT = 20


def _wf_parse_money(value: object) -> float | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    neg = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    cleaned = re.sub(r"[^0-9.]", "", s)
    if not cleaned:
        return None
    try:
        num = float(cleaned)
    except ValueError:
        return None
    return -num if neg else num


def _wf_parse_date(value: object):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s:
        return None
    try:
        return pd.to_datetime(s).to_pydatetime()
    except Exception:
        return None


def _wf_read_raw(raw_bytes: bytes, filename: str) -> pd.DataFrame:
    """Parse the Wells Fargo raw file (CSV or XLSX) into a clean DataFrame
    containing only real transaction rows."""
    name = filename.lower()
    if name.endswith(".xlsx") or name.endswith(".xlsm"):
        df = pd.read_excel(io.BytesIO(raw_bytes), header=None, dtype=str)
        raw_rows = [
            [
                "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
                for v in df.iloc[i].tolist()
            ]
            for i in range(len(df))
        ]
    else:
        text = raw_bytes.decode("utf-8-sig", errors="replace")
        raw_rows = [[c.strip() for c in row] for row in csv.reader(io.StringIO(text))]

    header_row_idx = None
    for i, row_vals in enumerate(raw_rows[:10]):
        if "Date" in row_vals and "Bank" in row_vals and "Net Charge" in row_vals:
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("Could not find the Wells Fargo header row (expected columns like 'Date', 'Bank', 'Net Charge').")

    header = raw_rows[header_row_idx]
    rows = []
    for row in raw_rows[header_row_idx + 1 :]:
        if len(row) < len(header):
            row = row + [""] * (len(header) - len(row))
        rows.append(row[: len(header)])
    data = pd.DataFrame(rows, columns=header)

    missing = [c for c in WF_OUT_COLS if c not in data.columns]
    if missing:
        raise ValueError(f"Raw file is missing expected columns: {', '.join(missing)}")

    data = data[WF_OUT_COLS].copy()

    # Drop footer rows like "Grand Total", "Total Funded", or fully-empty rows.
    def _is_data_row(row: pd.Series) -> bool:
        date_v = str(row["Date"]).strip() if row["Date"] is not None else ""
        bank_v = str(row["Bank"]).strip() if row["Bank"] is not None else ""
        if not date_v and not bank_v:
            return False
        if "total" in date_v.lower() or "total" in bank_v.lower():
            return False
        return True

    data = data[data.apply(_is_data_row, axis=1)].reset_index(drop=True)

    data["Date"] = data["Date"].map(_wf_parse_date)
    data["Tran Dt"] = data["Tran Dt"].map(_wf_parse_date)
    data["Gross-Chg"] = data["Gross-Chg"].map(_wf_parse_money)
    data["Discount"] = data["Discount"].map(_wf_parse_money)
    data["Net Charge"] = data["Net Charge"].map(_wf_parse_money)

    def _bank_key(b):
        s = str(b).strip()
        try:
            return (0, int(float(s)))
        except (TypeError, ValueError):
            return (1, s)

    data["__bank_key"] = data["Bank"].map(_bank_key)
    data["__tran_key"] = data["Tran Dt"].map(lambda d: d or datetime.max)
    data = data.sort_values(["__bank_key", "__tran_key"], kind="stable").drop(
        columns=["__bank_key", "__tran_key"]
    )
    return data.reset_index(drop=True)


def _wf_derive_filename(df: pd.DataFrame) -> str:
    for v in df["Date"]:
        if isinstance(v, datetime):
            return f"Formatted Wells Fargo - {v.strftime('%m-%d-%Y')}.xlsx"
    return f"Formatted Wells Fargo - {datetime.now().strftime('%m-%d-%Y')}.xlsx"


def write_wells_fargo_workbook(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Formatted"

    for col_idx, name in enumerate(WF_OUT_COLS, start=1):
        ws.cell(row=1, column=col_idx, value=name)

    cur = 2
    for bank_value, group in df.groupby("Bank", sort=False, dropna=False):
        group_start = cur
        for _, row in group.iterrows():
            ws.cell(row=cur, column=1, value=row["Date"]).number_format = WF_DATE_FMT
            ws.cell(row=cur, column=2, value=row["Tran Dt"]).number_format = WF_DATE_FMT
            bank_cell = ws.cell(row=cur, column=3)
            try:
                bank_cell.value = int(float(str(bank_value).strip()))
            except (TypeError, ValueError):
                bank_cell.value = bank_value
            ws.cell(row=cur, column=4, value=row["Acct"])
            ws.cell(row=cur, column=5, value=row["Name"])
            ws.cell(row=cur, column=6, value=row["Ticket"])
            for c_idx, col_name in ((7, "Gross-Chg"), (8, "Discount"), (9, "Net Charge")):
                cell = ws.cell(row=cur, column=c_idx, value=row[col_name])
                cell.number_format = WF_MONEY_FMT
            cur += 1
        group_end = cur - 1
        if group_end >= group_start:
            for c_idx, col_letter in ((7, "G"), (8, "H"), (9, "I")):
                cell = ws.cell(
                    row=cur,
                    column=c_idx,
                    value=f"=SUM({col_letter}{group_start}:{col_letter}{group_end})",
                )
                cell.number_format = WF_MONEY_FMT
                cell.font = WF_SUM_FONT
            ws.row_dimensions[cur].height = WF_SUM_ROW_HEIGHT
            cur += 1

    for col_letter, width in WF_COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _render_cardconnect_section() -> None:
    st.header("CardConnect Reformatter")
    st.caption(
        "Drop the daily raw credit-card transaction workbook below. "
        "The formatted BankReq file will download automatically."
    )

    uploaded = st.file_uploader(
        "Drag and drop the raw CardConnect .xlsx file here, or click to browse",
        type=["xlsx"],
        accept_multiple_files=False,
        key="cardconnect_upload",
    )
    if uploaded is None:
        return

    raw_bytes = uploaded.getvalue()
    try:
        sheets = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None)
    except Exception as e:
        st.error(f"Could not read the uploaded workbook: {e}")
        return

    branches = build_branches(sheets)
    if not branches:
        st.error(
            "No usable rows found. Expected a 'Net Sales' sheet with the "
            "standard BankReq columns."
        )
        return

    output_bytes = write_workbook(raw_bytes, branches)
    filename = derive_output_filename(branches)

    token = "cc-" + hashlib.sha1(raw_bytes).hexdigest()
    trigger_browser_download(output_bytes, filename, token)

    st.success(f"Formatted file ready: **{filename}**")
    st.download_button(
        "Download again",
        data=output_bytes,
        file_name=filename,
        mime=XLSX_MIME,
        key="cardconnect_download",
    )

    preview_rows = []
    for branch in branches:
        for product, amount in branch["net"]:
            preview_rows.append(
                {
                    "Site Alternate ID": branch["Site Alternate ID"],
                    "Funded Date": branch["Funded Date"],
                    "Site Name": branch["Site Name"],
                    "Product Code": product,
                    "Processed Transaction Amount": amount,
                    "Source": "Net Sales",
                }
            )
        for product, amount in branch["adj"]:
            preview_rows.append(
                {
                    "Site Alternate ID": branch["Site Alternate ID"],
                    "Funded Date": branch["Funded Date"],
                    "Site Name": branch["Site Name"],
                    "Product Code": product,
                    "Processed Transaction Amount": amount,
                    "Source": "Adjustments",
                }
            )
        for product, amount in branch["cb"]:
            preview_rows.append(
                {
                    "Site Alternate ID": branch["Site Alternate ID"],
                    "Funded Date": branch["Funded Date"],
                    "Site Name": branch["Site Name"],
                    "Product Code": product,
                    "Processed Transaction Amount": amount,
                    "Source": "Chargebacks",
                }
            )
    if preview_rows:
        st.subheader("Preview")
        preview_df = pd.DataFrame(preview_rows)

        def _highlight_red(row):
            return [
                "color: red" if row["Source"] in ("Adjustments", "Chargebacks") else ""
                for _ in row
            ]

        st.dataframe(
            preview_df.style.apply(_highlight_red, axis=1),
            use_container_width=True,
            hide_index=True,
        )


def _render_wells_fargo_section() -> None:
    st.header("Wells Fargo Funding Reformatter")
    st.caption(
        "Drop the raw Wells Fargo Funding file (.csv or .xlsx). "
        "The formatted workbook will download automatically."
    )

    uploaded = st.file_uploader(
        "Drag and drop the raw Wells Fargo file here, or click to browse",
        type=["csv", "xlsx"],
        accept_multiple_files=False,
        key="wf_upload",
    )
    if uploaded is None:
        return

    raw_bytes = uploaded.getvalue()
    try:
        wf_df = _wf_read_raw(raw_bytes, uploaded.name)
    except Exception as e:
        st.error(f"Could not read the Wells Fargo file: {e}")
        return

    if wf_df.empty:
        st.error("No transaction rows were found in the Wells Fargo file.")
        return

    output_bytes = write_wells_fargo_workbook(wf_df)
    filename = _wf_derive_filename(wf_df)

    token = "wf-" + hashlib.sha1(raw_bytes).hexdigest()
    trigger_browser_download(output_bytes, filename, token)

    st.success(f"Formatted file ready: **{filename}**")
    st.download_button(
        "Download again",
        data=output_bytes,
        file_name=filename,
        mime=XLSX_MIME,
        key="wf_download",
    )

    st.subheader("Preview")
    preview = wf_df.copy()
    preview["Date"] = preview["Date"].map(
        lambda d: d.strftime("%m-%d-%y") if isinstance(d, datetime) else d
    )
    preview["Tran Dt"] = preview["Tran Dt"].map(
        lambda d: d.strftime("%m-%d-%y") if isinstance(d, datetime) else d
    )
    st.dataframe(preview, use_container_width=True, hide_index=True)


def main() -> None:
    col_logo, col_title = st.columns([1, 5], vertical_alignment="center")
    with col_logo:
        if LOGO_FILE.exists():
            st.image(str(LOGO_FILE), width=160)
    with col_title:
        st.title("Stine BankReq Reformatter")
        st.caption("Use the section that matches the file you need to reformat.")

    _render_cardconnect_section()
    st.divider()
    _render_wells_fargo_section()


if __name__ == "__main__":
    main()
