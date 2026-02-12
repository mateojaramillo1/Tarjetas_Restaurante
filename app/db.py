import os
from datetime import datetime, timezone
from typing import Optional, List, Dict

import aiosqlite

DB_PATH = os.environ.get("DB_PATH", os.path.join("data", "card_reads.db"))


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
                read_at TEXT NOT NULL,
                FOREIGN KEY(person_id) REFERENCES people(id)
            )
            """
        )
        await db.execute(
            "CREATE INDEX IF NOT EXISTS idx_attendance_read_at ON attendance(read_at)"
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
        await db.execute(
            "INSERT INTO attendance (person_id, uid, atr, read_at) VALUES (?, ?, ?, ?)",
            (person["id"], uid, atr, read_at),
        )
        await db.commit()
    return {
        "uid": person["uid"],
        "first_name": person["first_name"],
        "last_name": person["last_name"],
        "id_number": person["id_number"],
        "phone": person["phone"],
        "area": person["area"],
        "atr": atr,
        "read_at": read_at,
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
