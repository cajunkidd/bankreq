"""Stine BankReq Reformatter — drag-and-drop the daily raw transaction file and
get back the reformatted BankReq workbook with a new Formatted sheet."""
from __future__ import annotations

import base64
import hashlib
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook

ROOT = Path(__file__).parent
LOGO_FILE = ROOT / "Stinelogo_white_rec.svg"

st.set_page_config(page_title="Stine BankReq Reformatter", layout="wide")

OUT_COLS = [
    "Site Alternate ID",
    "Funded Date",
    "Site Name",
    "Product Code",
    "Processed Transaction Amount",
]

# Product order matches the accounting team's template: Amex, DebitCard,
# Discover, Mastercard, Visa. Anything outside this list is appended in
# alphabetical order at the end of each site's group.
PRODUCT_ORDER = ["Amex", "DebitCard", "Discover", "Mastercard", "Visa"]

NET_SALES_SHEET = "Net Sales"
APPENDED_SHEETS = ["Adjustments", "Chargebacks & Chargeback Revers"]

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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


def build_sections(sheets: dict[str, pd.DataFrame]) -> list[tuple[str, pd.DataFrame]]:
    sections: list[tuple[str, pd.DataFrame]] = []
    net = sheets.get(NET_SALES_SHEET)
    if net is not None and not net.empty:
        block = _group_section(net)
        if not block.empty:
            sections.append((NET_SALES_SHEET, block))
    for name in APPENDED_SHEETS:
        df = sheets.get(name)
        if df is None or df.empty:
            continue
        block = _group_section(df)
        if not block.empty:
            sections.append((name, block))
    return sections


def derive_output_filename(sections: list[tuple[str, pd.DataFrame]]) -> str:
    for _, df in sections:
        if df.empty:
            continue
        raw = df.iloc[0]["Funded Date"]
        try:
            dt = pd.to_datetime(raw)
            return f"BankReq - {dt.strftime('%m-%d-%Y')}.xlsx"
        except Exception:
            break
    return f"BankReq - {datetime.now().strftime('%m-%d-%Y')}.xlsx"


def write_workbook(raw_bytes: bytes, sections: list[tuple[str, pd.DataFrame]]) -> bytes:
    wb = load_workbook(io.BytesIO(raw_bytes))
    if "Formatted" in wb.sheetnames:
        del wb["Formatted"]
    ws = wb.create_sheet("Formatted", 0)

    ws.append(OUT_COLS)
    first = True
    for label, df in sections:
        if first:
            first = False
        else:
            ws.append([])
            ws.append([f"From: {label}"])
        for site_id, funded, site_name, product, amount in df.itertuples(index=False):
            ws.append(
                [
                    site_id,
                    funded,
                    site_name,
                    product,
                    float(amount) if pd.notna(amount) else None,
                ]
            )

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def trigger_browser_download(data: bytes, filename: str, token: str) -> None:
    """Auto-download via a hidden anchor click. Keyed on `token` so a given
    upload only auto-downloads once per session even as Streamlit reruns."""
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


def main() -> None:
    col_logo, col_title = st.columns([1, 5], vertical_alignment="center")
    with col_logo:
        if LOGO_FILE.exists():
            st.image(str(LOGO_FILE), width=160)
    with col_title:
        st.title("BankReq Reformatter")
        st.caption(
            "Drop the daily raw credit-card transaction workbook below. "
            "The formatted BankReq file will download automatically."
        )

    uploaded = st.file_uploader(
        "Drag and drop the raw .xlsx file here, or click to browse",
        type=["xlsx"],
        accept_multiple_files=False,
    )
    if uploaded is None:
        return

    raw_bytes = uploaded.getvalue()
    try:
        sheets = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None)
    except Exception as e:
        st.error(f"Could not read the uploaded workbook: {e}")
        return

    sections = build_sections(sheets)
    if not sections:
        st.error(
            "No usable rows found. Expected a 'Net Sales' sheet with the "
            "standard BankReq columns."
        )
        return

    output_bytes = write_workbook(raw_bytes, sections)
    filename = derive_output_filename(sections)

    token = hashlib.sha1(raw_bytes).hexdigest()
    trigger_browser_download(output_bytes, filename, token)

    st.success(f"Formatted file ready: **{filename}**")
    st.download_button(
        "Download again",
        data=output_bytes,
        file_name=filename,
        mime=XLSX_MIME,
    )

    st.subheader("Preview")
    for label, df in sections:
        if label != NET_SALES_SHEET:
            st.markdown(f"**From: {label}**")
        st.dataframe(df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
