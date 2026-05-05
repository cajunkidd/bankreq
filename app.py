"""Stine Bank Data Viewer — desktop GUI (Tkinter).

Bundled to a single .exe via PyInstaller; data and logo are packed inside.
"""
from __future__ import annotations

import io
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd

APP_NAME = "Bank Data Viewer"


def resource_path(name: str) -> Path:
    base = getattr(sys, "_MEIPASS", None) or str(Path(__file__).resolve().parent)
    return Path(base) / name


DATA_FILE = resource_path("raw data (2).xlsx")
LOGO_FILE = resource_path("Stinelogo.png")
DATE_FORMAT = "%m/%d/%Y"


def _coerce_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Excel sometimes stores dates as text. For any column whose name contains
    # "date", try parsing as MM/DD/YYYY; adopt the parsed Timestamps only if
    # at least 80% of non-empty values parsed cleanly. Keeps display identical
    # while enabling chronological sort and date-aware filtering.
    for col in df.columns:
        if "date" not in str(col).lower():
            continue
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            continue
        parsed = pd.to_datetime(df[col], format=DATE_FORMAT, errors="coerce")
        non_empty = df[col].notna().sum()
        if non_empty and parsed.notna().sum() >= non_empty * 0.8:
            df[col] = parsed
    return df


def load_workbook(path: Path) -> dict[str, pd.DataFrame]:
    sheets = pd.read_excel(path, sheet_name=None)
    return {
        name: _coerce_date_columns(df.dropna(axis=1, how="all"))
        for name, df in sheets.items()
        if not df.dropna(how="all").empty
    }


def fmt_cell(value) -> str:
    if value is pd.NaT or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime(DATE_FORMAT)
    if isinstance(value, float):
        if value.is_integer():
            return f"{int(value):,}"
        return f"{value:,.2f}"
    return str(value)


class App(tk.Tk):
    def __init__(self, sheets: dict[str, pd.DataFrame]) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x800")
        self.minsize(900, 600)

        self.sheets = sheets
        self.current_sheet: str = next(iter(sheets))
        self.df: pd.DataFrame = sheets[self.current_sheet]
        self.filtered: pd.DataFrame = self.df
        self.filter_listboxes: dict[str, tk.Listbox] = {}
        self.sort_state: tuple[str, bool] | None = None  # (column, ascending)

        self._build_style()
        self._build_layout()
        self._populate_filters()
        self._refresh_table()

    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use(style.theme_use())
        style.configure("Treeview", rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Metric.TLabel", font=("Segoe UI", 10))

    def _build_layout(self) -> None:
        header = ttk.Frame(self, padding=(12, 8))
        header.pack(side=tk.TOP, fill=tk.X)

        if LOGO_FILE.exists():
            try:
                self._logo_img = tk.PhotoImage(file=str(LOGO_FILE))
                # Scale down if very large.
                while self._logo_img.width() > 220:
                    self._logo_img = self._logo_img.subsample(2, 2)
                ttk.Label(header, image=self._logo_img).pack(side=tk.LEFT, padx=(0, 12))
            except tk.TclError:
                pass

        ttk.Label(header, text=APP_NAME, style="Header.TLabel").pack(side=tk.LEFT)

        toolbar = ttk.Frame(self, padding=(12, 0, 12, 8))
        toolbar.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(toolbar, text="Sheet:").pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar(value=self.current_sheet)
        sheet_box = ttk.Combobox(
            toolbar,
            textvariable=self.sheet_var,
            values=list(self.sheets.keys()),
            state="readonly",
            width=32,
        )
        sheet_box.pack(side=tk.LEFT, padx=(6, 16))
        sheet_box.bind("<<ComboboxSelected>>", self._on_sheet_change)

        ttk.Label(toolbar, text="Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(toolbar, textvariable=self.search_var, width=32).pack(
            side=tk.LEFT, padx=(6, 8)
        )
        ttk.Button(toolbar, text="Reset filters", command=self._reset_filters).pack(
            side=tk.LEFT
        )

        body = ttk.Frame(self, padding=(12, 0))
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Scrollable filter sidebar.
        sidebar_outer = ttk.LabelFrame(body, text="Filters", padding=6)
        sidebar_outer.pack(side=tk.LEFT, fill=tk.Y)

        canvas = tk.Canvas(sidebar_outer, width=240, highlightthickness=0)
        sb = ttk.Scrollbar(sidebar_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.Y, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.filter_frame = ttk.Frame(canvas)
        self.filter_frame_id = canvas.create_window(
            (0, 0), window=self.filter_frame, anchor="nw"
        )
        self.filter_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfigure(self.filter_frame_id, width=e.width),
        )
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"),
        )

        # Table area.
        table_wrap = ttk.Frame(body)
        table_wrap.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))

        self.tree = ttk.Treeview(table_wrap, show="headings")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        table_wrap.rowconfigure(0, weight=1)
        table_wrap.columnconfigure(0, weight=1)

        footer = ttk.Frame(self, padding=(12, 8))
        footer.pack(side=tk.BOTTOM, fill=tk.X)

        self.row_count_var = tk.StringVar()
        self.metric_var = tk.StringVar()
        ttk.Label(footer, textvariable=self.row_count_var, style="Metric.TLabel").pack(
            side=tk.LEFT
        )
        ttk.Label(
            footer, textvariable=self.metric_var, style="Metric.TLabel"
        ).pack(side=tk.LEFT, padx=(20, 0))

        ttk.Button(footer, text="Export XLSX", command=self._export_xlsx).pack(
            side=tk.RIGHT, padx=(8, 0)
        )
        ttk.Button(footer, text="Export CSV", command=self._export_csv).pack(
            side=tk.RIGHT
        )

    def _populate_filters(self) -> None:
        for child in self.filter_frame.winfo_children():
            child.destroy()
        self.filter_listboxes.clear()

        for col in self.df.columns:
            unique = self.df[col].dropna().unique().tolist()
            if not (2 <= len(unique) <= 50):
                continue
            try:
                unique_sorted = sorted(unique, key=str)
            except TypeError:
                unique_sorted = unique

            frame = ttk.LabelFrame(self.filter_frame, text=str(col), padding=4)
            frame.pack(fill=tk.X, pady=4, padx=2)
            height = min(6, len(unique_sorted))
            lb = tk.Listbox(
                frame,
                selectmode=tk.EXTENDED,
                height=height,
                exportselection=False,
                activestyle="none",
            )
            for v in unique_sorted:
                lb.insert(tk.END, fmt_cell(v))
            lb.pack(fill=tk.X)
            lb._values = unique_sorted  # type: ignore[attr-defined]
            lb.bind("<<ListboxSelect>>", lambda *_: self._apply_filters())
            self.filter_listboxes[col] = lb

    def _on_sheet_change(self, *_args) -> None:
        self.current_sheet = self.sheet_var.get()
        self.df = self.sheets[self.current_sheet]
        self.search_var.set("")
        self.sort_state = None
        self._populate_filters()
        self._apply_filters()

    def _reset_filters(self) -> None:
        self.search_var.set("")
        for lb in self.filter_listboxes.values():
            lb.selection_clear(0, tk.END)
        self._apply_filters()

    def _apply_filters(self) -> None:
        out = self.df
        for col, lb in self.filter_listboxes.items():
            idxs = lb.curselection()
            if not idxs:
                continue
            chosen = [lb._values[i] for i in idxs]  # type: ignore[attr-defined]
            out = out[out[col].isin(chosen)]
        query = self.search_var.get().strip()
        if query:
            mask = pd.Series(False, index=out.index)
            for col in out.columns:
                if pd.api.types.is_datetime64_any_dtype(out[col]):
                    series = out[col].dt.strftime(DATE_FORMAT).fillna("")
                else:
                    series = out[col].astype(str)
                mask = mask | series.str.contains(
                    query, case=False, na=False, regex=False
                )
            out = out[mask]
        self.filtered = out
        if self.sort_state is not None:
            col, asc = self.sort_state
            if col in self.filtered.columns:
                self.filtered = self.filtered.sort_values(
                    col, ascending=asc, kind="mergesort"
                )
        self._refresh_table()

    def _refresh_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        cols = list(self.filtered.columns)
        self.tree.configure(columns=cols)
        for col in cols:
            self.tree.heading(
                col, text=str(col), command=lambda c=col: self._sort_by(c)
            )
            sample_len = max(
                [len(str(col))]
                + [len(fmt_cell(v)) for v in self.filtered[col].head(20)]
            )
            width = min(280, max(80, sample_len * 9))
            anchor = (
                tk.E if pd.api.types.is_numeric_dtype(self.filtered[col]) else tk.W
            )
            self.tree.column(col, width=width, anchor=anchor, stretch=False)

        for _, row in self.filtered.iterrows():
            self.tree.insert(
                "", tk.END, values=[fmt_cell(row[c]) for c in cols]
            )

        total = len(self.df)
        shown = len(self.filtered)
        self.row_count_var.set(f"Rows: {shown:,} of {total:,}")

        amt_col = next((c for c in cols if "Amount" in str(c)), None)
        if amt_col and pd.api.types.is_numeric_dtype(self.filtered[amt_col]):
            self.metric_var.set(
                f"Sum of {amt_col}: {self.filtered[amt_col].sum():,.2f}"
            )
        else:
            self.metric_var.set("")

    def _sort_by(self, col: str) -> None:
        asc = True
        if self.sort_state and self.sort_state[0] == col:
            asc = not self.sort_state[1]
        self.sort_state = (col, asc)
        self._apply_filters()

    def _export_frame(self) -> pd.DataFrame:
        # Render datetime columns back to MM/DD/YYYY so exports match the
        # source file format users expect.
        out = self.filtered.copy()
        for col in out.columns:
            if pd.api.types.is_datetime64_any_dtype(out[col]):
                out[col] = out[col].dt.strftime(DATE_FORMAT).where(
                    out[col].notna(), ""
                )
        return out

    def _export_csv(self) -> None:
        if self.filtered.empty:
            messagebox.showinfo(APP_NAME, "Nothing to export — the current view is empty.")
            return
        path = filedialog.asksaveasfilename(
            title="Export filtered view to CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile=f"{self._safe_name()}.csv",
        )
        if not path:
            return
        try:
            self._export_frame().to_csv(path, index=False)
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Could not save CSV:\n{e}")
            return
        messagebox.showinfo(APP_NAME, f"Exported {len(self.filtered):,} rows to:\n{path}")

    def _export_xlsx(self) -> None:
        if self.filtered.empty:
            messagebox.showinfo(APP_NAME, "Nothing to export — the current view is empty.")
            return
        path = filedialog.asksaveasfilename(
            title="Export filtered view to XLSX",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=f"{self._safe_name()}.xlsx",
        )
        if not path:
            return
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                self._export_frame().to_excel(
                    writer, sheet_name=self.current_sheet[:31] or "Sheet1", index=False
                )
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Could not save XLSX:\n{e}")
            return
        messagebox.showinfo(APP_NAME, f"Exported {len(self.filtered):,} rows to:\n{path}")

    def _safe_name(self) -> str:
        return f"{self.current_sheet.strip().replace(' ', '_')}_filtered"


def main() -> None:
    if not DATA_FILE.exists():
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            APP_NAME, f"Data file not found:\n{DATA_FILE}\n\nThis is a bundling problem."
        )
        return
    try:
        sheets = load_workbook(DATA_FILE)
    except Exception as e:  # noqa: BLE001 - surface any read error to user
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_NAME, f"Failed to read workbook:\n{e}")
        return
    if not sheets:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_NAME, "Workbook contains no populated sheets.")
        return
    App(sheets).mainloop()


if __name__ == "__main__":
    main()
