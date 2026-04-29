import sys
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from . import history, lumber, news, rates, weather
from .transform import reformat_workbook

app = FastAPI(title="Bank Reconciliation Reformatter")


def _static_dir() -> Path:
    # When frozen by PyInstaller, data files are extracted into _MEIPASS.
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "app" / "static"
    return Path(__file__).parent / "static"


STATIC_DIR = _static_dir()
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

INDEX_HTML = (STATIC_DIR / "index.html").read_text(encoding="utf-8")

XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return INDEX_HTML


@app.get("/api/weather")
def api_weather() -> JSONResponse:
    data = weather.get_forecast()
    if data is None:
        return JSONResponse({"available": False}, status_code=200)
    return JSONResponse({"available": True, **data})


@app.get("/api/news")
def api_news() -> JSONResponse:
    items = news.get_headlines()
    return JSONResponse({"items": items})


@app.get("/api/lumber")
def api_lumber() -> JSONResponse:
    q = lumber.get_quote()
    if q is None:
        return JSONResponse({"available": False})
    return JSONResponse({"available": True, **q})


@app.get("/api/rates")
def api_rates() -> JSONResponse:
    return JSONResponse({"items": rates.get_rates()})


@app.get("/api/uploads")
def api_uploads() -> JSONResponse:
    return JSONResponse({"items": history.list_recent_files(limit=10)})


@app.get("/api/uploads/{file_id}/download")
def api_uploads_download(file_id: int) -> Response:
    fetched = history.get_file(file_id)
    if fetched is None:
        raise HTTPException(404, "File not found.")
    file_bytes, filename = fetched
    return Response(
        content=file_bytes,
        media_type=XLSX_MIME,
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition",
        },
    )


@app.post("/reformat")
async def reformat(file: UploadFile = File(...)) -> Response:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Please upload an .xlsx file.")
    raw = await file.read()
    try:
        out_bytes, _sheet_name, funded_date, anomaly_count = reformat_workbook(raw)
    except ValueError as e:
        raise HTTPException(400, str(e))

    # MM-DD-YYYY (dashes — slashes aren't allowed in Windows filenames).
    out_name = f"BankReq - {funded_date.strftime('%m-%d-%Y')}.xlsx"
    history.record_formatted_file(
        out_bytes, out_name, funded_date, anomaly_count
    )
    return Response(
        content=out_bytes,
        media_type=XLSX_MIME,
        headers={
            "Content-Disposition": f'attachment; filename="{out_name}"',
            "X-Anomaly-Count": str(anomaly_count),
            "Access-Control-Expose-Headers": "Content-Disposition, X-Anomaly-Count",
        },
    )
