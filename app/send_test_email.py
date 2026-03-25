import asyncio
import datetime

from app.db import init_db, list_attendance_filtered
from app.main import (
    _build_attendance_workbook,
    _load_report_config,
    _parse_dt_param,
    _send_email_with_attachment,
)


async def main() -> None:
    await init_db()
    cfg = _load_report_config()
    today = datetime.date.today()
    from_raw = f"{today.isoformat()}T00:00:00"
    to_raw = f"{today.isoformat()}T23:59:59"

    rows = await list_attendance_filtered(
        from_dt=_parse_dt_param(from_raw),
        to_dt=_parse_dt_param(to_raw),
        name=None,
        id_number=None,
        area=None,
        uid=None,
        month_key=None,
        limit=50000,
    )
    rows = sorted(rows, key=lambda row: str(row.get("read_at") or ""))
    xlsx_data = _build_attendance_workbook(rows, sheet_name="Prueba correo")

    await asyncio.to_thread(
        _send_email_with_attachment,
        cfg,
        subject=f"Prueba envio control - {today.isoformat()}",
        body=(
            "Este es un correo de prueba del sistema de control. "
            f"Registros incluidos: {len(rows)}."
        ),
        filename=f"prueba-control-{today.isoformat()}.xlsx",
        attachment_bytes=xlsx_data,
    )
    print(f"OK: correo enviado. Registros en adjunto: {len(rows)}")


if __name__ == "__main__":
    asyncio.run(main())
