import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional, List, Dict

import aiosqlite

_PROJECT_ROOT = Path(__file__).resolve().parents[1]
DB_PATH = os.environ.get("DB_PATH", str(_PROJECT_ROOT / "data" / "card_reads.db"))
ATTENDANCE_COOLDOWN_HOURS = int(os.environ.get("ATTENDANCE_COOLDOWN_HOURS", "3"))


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


async def init_db() -> None:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            """
            CREATE TABLE IF NOT EXISTS people (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                uid TEXT NOT NULL UNIQUE,
                first_name TEXT NOT NULL,
                last_name TEXT NOT NULL,
                id_number TEXT NOT NULL,
                phone TEXT,
                area TEXT,
                created_at TEXT NOT NULL
            )
            """
        )
        await db.execute(
            """
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                person_id INTEGER NOT NULL,
                uid TEXT NOT NULL,
                atr TEXT,
                month_key TEXT,
                read_at TEXT NOT NULL,
                FOREIGN KEY(person_id) REFERENCES people(id)
            )
            """
        )
        cursor = await db.execute("PRAGMA table_info(attendance)")
        attendance_columns = await cursor.fetchall()
        await cursor.close()
        attendance_column_names = {column[1] for column in attendance_columns}
        if "month_key" not in attendance_column_names:
            await db.execute("ALTER TABLE attendance ADD COLUMN month_key TEXT")

        await db.execute(
            """
            UPDATE attendance
            SET month_key = substr(read_at, 1, 7)
            WHERE month_key IS NULL OR month_key = ''
            """
        )
        await db.execute(
            "CREATE INDEX IF NOT EXISTS idx_attendance_read_at ON attendance(read_at)"
        )
        await db.execute(
            "CREATE INDEX IF NOT EXISTS idx_attendance_month_key ON attendance(month_key)"
        )

        await db.execute("DROP TRIGGER IF EXISTS trg_attendance_cooldown")
        if ATTENDANCE_COOLDOWN_HOURS > 0:
            await db.execute(
                f"""
                CREATE TRIGGER trg_attendance_cooldown
                BEFORE INSERT ON attendance
                FOR EACH ROW
                WHEN EXISTS (
                    SELECT 1
                    FROM attendance a
                    WHERE a.person_id = NEW.person_id
                      AND julianday(NEW.read_at) >= julianday(a.read_at)
                      AND (julianday(NEW.read_at) - julianday(a.read_at)) * 24 < {ATTENDANCE_COOLDOWN_HOURS}
                )
                BEGIN
                    SELECT RAISE(ABORT, 'ATTENDANCE_COOLDOWN_ACTIVE');
                END;
                """
            )
        await db.commit()


async def upsert_person(
    uid: str,
    first_name: str,
    last_name: str,
    id_number: str,
    phone: Optional[str],
    area: Optional[str],
) -> Dict[str, str]:
    created_at = _now_iso()
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            """
            INSERT INTO people (uid, first_name, last_name, id_number, phone, area, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(uid) DO UPDATE SET
                first_name=excluded.first_name,
                last_name=excluded.last_name,
                id_number=excluded.id_number,
                phone=excluded.phone,
                area=excluded.area
            """,
            (uid, first_name, last_name, id_number, phone, area, created_at),
        )
        await db.commit()
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            "SELECT uid, first_name, last_name, id_number, phone, area, created_at FROM people WHERE uid = ?",
            (uid,),
        )
        row = await cursor.fetchone()
        await cursor.close()
    return {
        "uid": row["uid"],
        "first_name": row["first_name"],
        "last_name": row["last_name"],
        "id_number": row["id_number"],
        "phone": row["phone"],
        "area": row["area"],
        "created_at": row["created_at"],
    }


async def get_person_by_uid(uid: str) -> Optional[Dict[str, str]]:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            "SELECT uid, first_name, last_name, id_number, phone, area, created_at FROM people WHERE uid = ?",
            (uid,),
        )
        row = await cursor.fetchone()
        await cursor.close()
    if row is None:
        return None
    return {
        "uid": row["uid"],
        "first_name": row["first_name"],
        "last_name": row["last_name"],
        "id_number": row["id_number"],
        "phone": row["phone"],
        "area": row["area"],
        "created_at": row["created_at"],
    }


async def list_people(limit: int = 100) -> List[Dict[str, str]]:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            """
            SELECT uid, first_name, last_name, id_number, phone, area, created_at
            FROM people ORDER BY id DESC LIMIT ?
            """,
            (limit,),
        )
        rows = await cursor.fetchall()
        await cursor.close()
    return [
        {
            "uid": row["uid"],
            "first_name": row["first_name"],
            "last_name": row["last_name"],
            "id_number": row["id_number"],
            "phone": row["phone"],
            "area": row["area"],
            "created_at": row["created_at"],
        }
        for row in rows
    ]


async def list_people_all() -> List[Dict[str, str]]:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            """
            SELECT uid, first_name, last_name, id_number, phone, area, created_at
            FROM people ORDER BY id ASC
            """
        )
        rows = await cursor.fetchall()
        await cursor.close()
    return [
        {
            "uid": row["uid"],
            "first_name": row["first_name"],
            "last_name": row["last_name"],
            "id_number": row["id_number"],
            "phone": row["phone"],
            "area": row["area"],
            "created_at": row["created_at"],
        }
        for row in rows
    ]


async def list_people_filtered(
    from_dt: Optional[str],
    to_dt: Optional[str],
    name: Optional[str],
    id_number: Optional[str],
    area: Optional[str],
    limit: int = 200,
) -> List[Dict[str, str]]:
    query = [
        "SELECT uid, first_name, last_name, id_number, phone, area, created_at",
        "FROM people",
        "WHERE 1=1",
    ]
    params: List[str] = []
    if from_dt:
        query.append("AND created_at >= ?")
        params.append(from_dt)
    if to_dt:
        query.append("AND created_at <= ?")
        params.append(to_dt)
    if name:
        query.append("AND (lower(first_name) LIKE ? OR lower(last_name) LIKE ?)")
        like_value = f"%{name.lower()}%"
        params.extend([like_value, like_value])
    if id_number:
        query.append("AND lower(id_number) LIKE ?")
        params.append(f"%{id_number.lower()}%")
    if area:
        query.append("AND lower(area) LIKE ?")
        params.append(f"%{area.lower()}%")
    query.append("ORDER BY id DESC LIMIT ?")
    params.append(str(limit))

    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute("\n".join(query), tuple(params))
        rows = await cursor.fetchall()
        await cursor.close()
    return [
        {
            "uid": row["uid"],
            "first_name": row["first_name"],
            "last_name": row["last_name"],
            "id_number": row["id_number"],
            "phone": row["phone"],
            "area": row["area"],
            "created_at": row["created_at"],
        }
        for row in rows
    ]


async def add_attendance(uid: str, atr: Optional[str]) -> Optional[Dict[str, str]]:
    read_at = _now_iso()
    now_utc = datetime.now(timezone.utc)
    month_key = datetime.now().strftime("%Y-%m")
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            "SELECT id, uid, first_name, last_name, id_number, phone, area FROM people WHERE uid = ?",
            (uid,),
        )
        person = await cursor.fetchone()
        await cursor.close()
        if person is None:
            return None

        cursor = await db.execute(
            "SELECT read_at FROM attendance WHERE person_id = ? ORDER BY id DESC LIMIT 1",
            (person["id"],),
        )
        last_row = await cursor.fetchone()
        await cursor.close()

        if last_row is not None and ATTENDANCE_COOLDOWN_HOURS > 0:
            try:
                last_read_at = datetime.fromisoformat(last_row["read_at"])
                if last_read_at.tzinfo is None:
                    last_read_at = last_read_at.replace(tzinfo=timezone.utc)
                cooldown = timedelta(hours=ATTENDANCE_COOLDOWN_HOURS)
                elapsed = now_utc - last_read_at.astimezone(timezone.utc)
                if elapsed < cooldown:
                    allowed_at = last_read_at.astimezone(timezone.utc) + cooldown
                    return {
                        "uid": person["uid"],
                        "first_name": person["first_name"],
                        "last_name": person["last_name"],
                        "id_number": person["id_number"],
                        "phone": person["phone"],
                        "area": person["area"],
                        "atr": atr,
                        "read_at": last_row["read_at"],
                        "skipped": True,
                        "allowed_at": allowed_at.isoformat(),
                    }
            except Exception:
                pass

        try:
            await db.execute(
                "INSERT INTO attendance (person_id, uid, atr, month_key, read_at) VALUES (?, ?, ?, ?, ?)",
                (person["id"], uid, atr, month_key, read_at),
            )
            await db.commit()
        except aiosqlite.IntegrityError as exc:
            if "ATTENDANCE_COOLDOWN_ACTIVE" in str(exc):
                allowed_at = None
                if last_row is not None and ATTENDANCE_COOLDOWN_HOURS > 0:
                    try:
                        last_read_at = datetime.fromisoformat(last_row["read_at"])
                        if last_read_at.tzinfo is None:
                            last_read_at = last_read_at.replace(tzinfo=timezone.utc)
                        allowed_at = (last_read_at.astimezone(timezone.utc) + timedelta(hours=ATTENDANCE_COOLDOWN_HOURS)).isoformat()
                    except Exception:
                        allowed_at = None
                return {
                    "uid": person["uid"],
                    "first_name": person["first_name"],
                    "last_name": person["last_name"],
                    "id_number": person["id_number"],
                    "phone": person["phone"],
                    "area": person["area"],
                    "atr": atr,
                    "read_at": last_row["read_at"] if last_row else read_at,
                    "skipped": True,
                    "allowed_at": allowed_at,
                }
            raise
    return {
        "uid": person["uid"],
        "first_name": person["first_name"],
        "last_name": person["last_name"],
        "id_number": person["id_number"],
        "phone": person["phone"],
        "area": person["area"],
        "atr": atr,
        "read_at": read_at,
        "skipped": False,
    }


async def list_attendance(limit: int = 100) -> List[Dict[str, str]]:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            """
            SELECT a.uid, a.atr, a.read_at,
                   p.first_name, p.last_name, p.id_number, p.phone, p.area
            FROM attendance a
            JOIN people p ON p.id = a.person_id
            ORDER BY a.id DESC LIMIT ?
            """,
            (limit,),
        )
        rows = await cursor.fetchall()
        await cursor.close()
    return [
        {
            "uid": row["uid"],
            "first_name": row["first_name"],
            "last_name": row["last_name"],
            "id_number": row["id_number"],
            "phone": row["phone"],
            "area": row["area"],
            "atr": row["atr"],
            "read_at": row["read_at"],
        }
        for row in rows
    ]


async def list_attendance_filtered(
    from_dt: Optional[str],
    to_dt: Optional[str],
    name: Optional[str],
    id_number: Optional[str],
    area: Optional[str],
    uid: Optional[str],
    month_key: Optional[str],
    limit: int = 200,
) -> List[Dict[str, str]]:
    query = [
        "SELECT a.uid, a.atr, a.read_at,",
        "       p.first_name, p.last_name, p.id_number, p.phone, p.area",
        "FROM attendance a",
        "JOIN people p ON p.id = a.person_id",
        "WHERE 1=1",
    ]
    params: List[str] = []
    if from_dt:
        query.append("AND a.read_at >= ?")
        params.append(from_dt)
    if to_dt:
        query.append("AND a.read_at <= ?")
        params.append(to_dt)
    if name:
        query.append("AND (lower(p.first_name) LIKE ? OR lower(p.last_name) LIKE ?)")
        like_value = f"%{name.lower()}%"
        params.extend([like_value, like_value])
    if id_number:
        query.append("AND lower(p.id_number) LIKE ?")
        params.append(f"%{id_number.lower()}%")
    if area:
        query.append("AND lower(p.area) LIKE ?")
        params.append(f"%{area.lower()}%")
    if uid:
        query.append("AND lower(a.uid) LIKE ?")
        params.append(f"%{uid.lower()}%")
    if month_key:
        query.append("AND a.month_key = ?")
        params.append(month_key)
    query.append("ORDER BY a.id DESC LIMIT ?")
    params.append(str(limit))

    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute("\n".join(query), tuple(params))
        rows = await cursor.fetchall()
        await cursor.close()
    return [
        {
            "uid": row["uid"],
            "first_name": row["first_name"],
            "last_name": row["last_name"],
            "id_number": row["id_number"],
            "phone": row["phone"],
            "area": row["area"],
            "atr": row["atr"],
            "read_at": row["read_at"],
        }
        for row in rows
    ]


async def list_attendance_all() -> List[Dict[str, str]]:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute(
            """
            SELECT a.uid, a.atr, a.read_at,
                   p.first_name, p.last_name, p.id_number, p.phone, p.area
            FROM attendance a
            JOIN people p ON p.id = a.person_id
            ORDER BY a.id ASC
            """
        )
        rows = await cursor.fetchall()
        await cursor.close()
    return [
        {
            "uid": row["uid"],
            "first_name": row["first_name"],
            "last_name": row["last_name"],
            "id_number": row["id_number"],
            "phone": row["phone"],
            "area": row["area"],
            "atr": row["atr"],
            "read_at": row["read_at"],
        }
        for row in rows
    ]
