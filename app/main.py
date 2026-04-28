from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles

from .transform import reformat_workbook

app = FastAPI(title="Bank Reconciliation Reformatter")

STATIC_DIR = Path(__file__).parent / "static"
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

INDEX_HTML = (STATIC_DIR / "index.html").read_text()

XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return INDEX_HTML


@app.post("/reformat")
async def reformat(file: UploadFile = File(...)) -> Response:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Please upload an .xlsx file.")
    raw = await file.read()
    try:
        out_bytes, _sheet_name, funded_date = reformat_workbook(raw)
    except ValueError as e:
        raise HTTPException(400, str(e))

    # MM-DD-YYYY (dashes — slashes aren't allowed in Windows filenames).
    out_name = f"BankReq - {funded_date.strftime('%m-%d-%Y')}.xlsx"
    return Response(
        content=out_bytes,
        media_type=XLSX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )
