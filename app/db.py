"""SQLite 连接与建表。"""
import sqlite3
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS students (
    student_id  TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    major       TEXT,
    class_name  TEXT,
    grade       TEXT,
    extra_json  TEXT
);

CREATE TABLE IF NOT EXISTS absence_records (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id  TEXT NOT NULL,
    week        INTEGER NOT NULL,
    weekday     TEXT NOT NULL DEFAULT '',
    period      TEXT NOT NULL,
    type        TEXT NOT NULL CHECK (type IN ('absent','leave')),
    note        TEXT,
    created_at  TEXT NOT NULL,
    UNIQUE (student_id, week, weekday, period)
);
"""

OLD_COLUMNS = ["id", "student_id", "week", "period", "type", "note", "created_at"]


def connect(db_path: Path) -> sqlite3.Connection:
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _migrate_add_weekday(conn: sqlite3.Connection) -> None:
    """旧表（无 weekday 列）重建为新结构，历史记录 weekday 置空。"""
    conn.execute("ALTER TABLE absence_records RENAME TO absence_records_old")
    conn.executescript(SCHEMA)
    conn.execute(
        "INSERT INTO absence_records (id, student_id, week, weekday, period, type, note, created_at)"
        " SELECT id, student_id, week, '', period, type, note, created_at"
        " FROM absence_records_old")
    conn.execute("DROP TABLE absence_records_old")


def init(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA)
    cols = [r[1] for r in conn.execute("PRAGMA table_info(absence_records)")]
    if "weekday" not in cols:
        _migrate_add_weekday(conn)
    conn.commit()
