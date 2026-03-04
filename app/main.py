import asyncio
import os
import re
import time
try:
    import winsound
except Exception:
    winsound = None
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Optional, Dict

from fastapi import FastAPI, HTTPException, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware
from pydantic import BaseModel, Field
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from .db import (
    init_db,
    upsert_person,
    get_person_by_uid,
    list_people,
    list_people_filtered,
    list_people_all,
    add_attendance,
    list_attendance,
    list_attendance_filtered,
    list_attendance_all,
)
from .reader import CardReaderService, CardRead

app = FastAPI(title="Lector Omnikey", version="0.1.0")
app.add_middleware(SessionMiddleware, secret_key=os.environ.get("SESSION_SECRET", "change-me"))

static_dir = Path(__file__).parent / "static"
app.mount("/static", StaticFiles(directory=static_dir), name="static")

_reader_service: Optional[CardReaderService] = None
_loop: Optional[asyncio.AbstractEventLoop] = None
_last_read: Optional[Dict[str, str]] = None
_reader_state: Dict[str, Optional[str]] = {"present": False, "last_seen_at": None}
_control_last_ping_monotonic: float = 0.0
_control_ping_ttl_seconds: float = 4.0


class PersonIn(BaseModel):
    uid: str = Field(..., min_length=4)
    first_name: str = Field(..., min_length=1)
    last_name: str = Field(..., min_length=1)
    id_number: str = Field(..., min_length=3)
    phone: Optional[str] = None
    area: Optional[str] = None


def _is_admin(request: Request) -> bool:
    return bool(request.session.get("admin"))


def _require_admin(request: Request) -> None:
    if not _is_admin(request):
        raise HTTPException(status_code=401, detail="Admin login required")


def _autosize_columns(ws) -> None:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)


def _format_dt(value: str) -> str:
    try:
        dt = datetime.fromisoformat(value)
        return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return value


def _normalize_query(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    value = value.strip()
    return value or None


def _parse_dt_param(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value)
        if dt.tzinfo is None:
            local_tz = datetime.now().astimezone().tzinfo
            dt = dt.replace(tzinfo=local_tz)
        return dt.astimezone(timezone.utc).isoformat()
    except Exception:
        return None


def _parse_month_param(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    value = value.strip()
    if re.fullmatch(r"\d{4}-\d{2}", value):
        return value
    return None


@app.on_event("startup")
async def on_startup() -> None:
    global _reader_service, _loop
    await init_db()
    _loop = asyncio.get_running_loop()

    def handle_card(read: CardRead) -> None:
        global _last_read
        if _loop is None:
            return
        read_at = datetime.now(timezone.utc).isoformat()
        _last_read = {"uid": read.uid, "atr": read.atr, "read_at": read_at}
        _reader_state["present"] = True
        _reader_state["last_seen_at"] = read_at
        if winsound is not None:
            try:
                winsound.MessageBeep(winsound.MB_OK)
            except Exception:
                pass
        print(f"UID leido: {read.uid} | ATR: {read.atr}")
        control_recently_active = (time.monotonic() - _control_last_ping_monotonic) <= _control_ping_ttl_seconds
        if control_recently_active:
            future = asyncio.run_coroutine_threadsafe(add_attendance(read.uid, read.atr), _loop)
            future.add_done_callback(lambda _: None)

    def handle_remove() -> None:
        _reader_state["present"] = False

    _reader_service = CardReaderService(on_card=handle_card, on_remove=handle_remove)
    _reader_service.start()


@app.on_event("shutdown")
async def on_shutdown() -> None:
    if _reader_service:
        _reader_service.stop()


@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    html_path = static_dir / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.get("/admin/login", response_class=HTMLResponse)
async def admin_login_page() -> HTMLResponse:
    html_path = static_dir / "login.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.post("/admin/login")
async def admin_login(request: Request, password: str = Form(...)) -> dict:
    admin_password = os.environ.get("ADMIN_PASSWORD", "admin123")
    if password != admin_password:
        raise HTTPException(status_code=401, detail="Clave incorrecta")
    request.session["admin"] = True
    return {"ok": True, "redirect": "/admin"}


@app.post("/admin/logout")
async def admin_logout(request: Request) -> RedirectResponse:
    request.session.clear()
    return RedirectResponse(url="/", status_code=303)


@app.get("/admin", response_class=HTMLResponse)
async def admin_page(request: Request) -> HTMLResponse:
    _require_admin(request)
    html_path = static_dir / "admin.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.get("/control", response_class=HTMLResponse)
async def control_page() -> HTMLResponse:
    html_path = static_dir / "control.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.post("/api/control/ping")
async def control_ping() -> dict:
    global _control_last_ping_monotonic
    _control_last_ping_monotonic = time.monotonic()
    return {"ok": True}


@app.get("/api/health")
async def health() -> dict:
    status = "ok"
    details = {}
    if _reader_service and _reader_service.init_error:
        status = "error"
        details["reader_error"] = _reader_service.init_error
    return {"status": status, "details": details}


@app.get("/api/latest")
async def latest() -> dict:
    if _last_read is None:
        return {"latest": None}
    person = await get_person_by_uid(_last_read["uid"])
    latest_payload = {**_last_read, "person": person, **_reader_state}
    return {"latest": latest_payload}


@app.post("/api/people")
async def create_person(payload: PersonIn, request: Request) -> dict:
    _require_admin(request)
    person = await upsert_person(
        uid=payload.uid,
        first_name=payload.first_name,
        last_name=payload.last_name,
        id_number=payload.id_number,
        phone=payload.phone,
        area=payload.area,
    )
    return {"person": person}


@app.get("/api/people")
async def people(request: Request, limit: int = 100) -> dict:
    _require_admin(request)
    data = await list_people(limit=limit)
    return {"people": data}


@app.get("/api/people/search")
async def people_search(
    request: Request,
    from_dt: Optional[str] = None,
    to_dt: Optional[str] = None,
    name: Optional[str] = None,
    id_number: Optional[str] = None,
    area: Optional[str] = None,
    limit: int = 200,
) -> dict:
    _require_admin(request)
    data = await list_people_filtered(
        from_dt=_parse_dt_param(from_dt),
        to_dt=_parse_dt_param(to_dt),
        name=_normalize_query(name),
        id_number=_normalize_query(id_number),
        area=_normalize_query(area),
        limit=limit,
    )
    return {"people": data}


@app.get("/api/people/{uid}")
async def person_by_uid(uid: str, request: Request) -> dict:
    _require_admin(request)
    person = await get_person_by_uid(uid)
    if person is None:
        raise HTTPException(status_code=404, detail="Person not found")
    return {"person": person}


@app.get("/api/attendance")
async def attendance(
    request: Request,
    from_dt: Optional[str] = None,
    to_dt: Optional[str] = None,
    name: Optional[str] = None,
    id_number: Optional[str] = None,
    area: Optional[str] = None,
    uid: Optional[str] = None,
    month: Optional[str] = None,
    limit: int = 200,
) -> dict:
    _require_admin(request)
    data = await list_attendance_filtered(
        from_dt=_parse_dt_param(from_dt),
        to_dt=_parse_dt_param(to_dt),
        name=_normalize_query(name),
        id_number=_normalize_query(id_number),
        area=_normalize_query(area),
        uid=_normalize_query(uid),
        month_key=_parse_month_param(month),
        limit=limit,
    )
    return {"attendance": data}


@app.get("/api/export.xlsx")
async def export_xlsx(
    request: Request,
    from_dt: Optional[str] = None,
    to_dt: Optional[str] = None,
    name: Optional[str] = None,
    id_number: Optional[str] = None,
    area: Optional[str] = None,
    uid: Optional[str] = None,
    month: Optional[str] = None,
) -> StreamingResponse:
    _require_admin(request)
    rows = await list_attendance_filtered(
        from_dt=_parse_dt_param(from_dt),
        to_dt=_parse_dt_param(to_dt),
        name=_normalize_query(name),
        id_number=_normalize_query(id_number),
        area=_normalize_query(area),
        uid=_normalize_query(uid),
        month_key=_parse_month_param(month),
        limit=50000,
    )
    wb = Workbook()
    ws = wb.active
    ws.title = "Control"
    headers = [
        "UID",
        "Nombre",
        "Apellido",
        "Cedula",
        "Telefono",
        "Area",
        "ATR",
        "FechaHora",
    ]
    ws.append(headers)
    for row in rows:
        ws.append(
            [
                row["uid"],
                row["first_name"],
                row["last_name"],
                row["id_number"],
                row["phone"],
                row["area"],
                row["atr"],
                _format_dt(row["read_at"]),
            ]
        )
    ws.freeze_panes = "A2"
    table = Table(displayName="ControlTable", ref=f"A1:H{ws.max_row}")
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style
    ws.add_table(table)
    _autosize_columns(ws)
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    headers = {"Content-Disposition": "attachment; filename=control.xlsx"}
    return StreamingResponse(
        stream,
        headers=headers,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/export-people.xlsx")
async def export_people_xlsx(request: Request) -> StreamingResponse:
    _require_admin(request)
    rows = await list_people_all()
    wb = Workbook()
    ws = wb.active
    ws.title = "Personas"
    headers = ["UID", "Nombre", "Apellido", "Cedula", "Telefono", "Area", "Creado"]
    ws.append(headers)
    for row in rows:
        ws.append(
            [
                row["uid"],
                row["first_name"],
                row["last_name"],
                row["id_number"],
                row["phone"],
                row["area"],
                _format_dt(row["created_at"]),
            ]
        )
    ws.freeze_panes = "A2"
    table = Table(displayName="PersonasTable", ref=f"A1:G{ws.max_row}")
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style
    ws.add_table(table)
    _autosize_columns(ws)
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    headers = {"Content-Disposition": "attachment; filename=personas.xlsx"}
    return StreamingResponse(
        stream,
        headers=headers,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/export-people-filtered.xlsx")
async def export_people_filtered_xlsx(
    request: Request,
    from_dt: Optional[str] = None,
    to_dt: Optional[str] = None,
    name: Optional[str] = None,
    id_number: Optional[str] = None,
    area: Optional[str] = None,
) -> StreamingResponse:
    _require_admin(request)
    rows = await list_people_filtered(
        from_dt=_parse_dt_param(from_dt),
        to_dt=_parse_dt_param(to_dt),
        name=_normalize_query(name),
        id_number=_normalize_query(id_number),
        area=_normalize_query(area),
        limit=2000,
    )
    wb = Workbook()
    ws = wb.active
    ws.title = "Personas"
    headers = ["UID", "Nombre", "Apellido", "Cedula", "Telefono", "Area", "Creado"]
    ws.append(headers)
    for row in rows:
        ws.append(
            [
                row["uid"],
                row["first_name"],
                row["last_name"],
                row["id_number"],
                row["phone"],
                row["area"],
                _format_dt(row["created_at"]),
            ]
        )
    ws.freeze_panes = "A2"
    table = Table(displayName="PersonasFiltradas", ref=f"A1:G{ws.max_row}")
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style
    ws.add_table(table)
    _autosize_columns(ws)
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    headers = {"Content-Disposition": "attachment; filename=personas-filtradas.xlsx"}
    return StreamingResponse(
        stream,
        headers=headers,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
