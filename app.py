"""Stine Bank Data Viewer — local Streamlit app."""
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).parent
DATA_FILE = ROOT / "raw data (2).xlsx"
LOGO_FILE = ROOT / "Stinelogo_white_rec.svg"

st.set_page_config(page_title="Stine Bank Data Viewer", layout="wide")


@st.cache_data(show_spinner=False)
def load_workbook(path: Path) -> dict[str, pd.DataFrame]:
    sheets = pd.read_excel(path, sheet_name=None)
    return {name: df.dropna(axis=1, how="all") for name, df in sheets.items()}


def filter_df(df: pd.DataFrame, query: str, column_filters: dict[str, list]) -> pd.DataFrame:
    out = df
    for col, selected in column_filters.items():
        if selected:
            out = out[out[col].isin(selected)]
    if query:
        mask = out.apply(
            lambda row: row.astype(str).str.contains(query, case=False, na=False).any(),
            axis=1,
        )
        out = out[mask]
    return out


def to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buf = io.BytesIO()
    safe_name = sheet_name[:31] or "Sheet1"
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=safe_name, index=False)
    return buf.getvalue()


def main() -> None:
    col_logo, col_title = st.columns([1, 5], vertical_alignment="center")
    with col_logo:
        if LOGO_FILE.exists():
            st.image(str(LOGO_FILE), width=160)
    with col_title:
        st.title("Bank Data Viewer")
        st.caption(f"Source: {DATA_FILE.name}")

    if not DATA_FILE.exists():
        st.error(f"Data file not found: {DATA_FILE}")
        st.stop()

    sheets = load_workbook(DATA_FILE)
    populated = {name: df for name, df in sheets.items() if not df.empty}
    if not populated:
        st.warning("No populated sheets found in the workbook.")
        st.stop()

    with st.sidebar:
        st.header("Filters")
        sheet_name = st.selectbox("Sheet", list(populated.keys()))
        df = populated[sheet_name]

        query = st.text_input("Search (any column)", "")

        column_filters: dict[str, list] = {}
        candidate_cols = [
            c for c in df.columns
            if df[c].dtype == object or 2 <= df[c].nunique(dropna=True) <= 50
        ]
        for col in candidate_cols:
            options = sorted(df[col].dropna().unique().tolist(), key=str)
            if 1 < len(options) <= 50:
                column_filters[col] = st.multiselect(col, options, default=[])

    filtered = filter_df(df, query, column_filters)

    m1, m2, m3 = st.columns(3)
    m1.metric("Rows (filtered)", f"{len(filtered):,}")
    m2.metric("Rows (total)", f"{len(df):,}")
    amt_col = next((c for c in filtered.columns if "Amount" in c), None)
    if amt_col is not None and pd.api.types.is_numeric_dtype(filtered[amt_col]):
        m3.metric(f"Sum of {amt_col}", f"{filtered[amt_col].sum():,.2f}")

    st.dataframe(filtered, use_container_width=True, hide_index=True)

    st.subheader("Export filtered view")
    c1, c2 = st.columns(2)
    base = f"{sheet_name.strip().replace(' ', '_')}_filtered"
    with c1:
        st.download_button(
            "Download CSV",
            data=filtered.to_csv(index=False).encode("utf-8"),
            file_name=f"{base}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "Download XLSX",
            data=to_xlsx_bytes(filtered, sheet_name),
            file_name=f"{base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
