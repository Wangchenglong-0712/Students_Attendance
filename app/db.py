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
    period      TEXT NOT NULL,
    type        TEXT NOT NULL CHECK (type IN ('absent','leave')),
    note        TEXT,
    created_at  TEXT NOT NULL,
    UNIQUE (student_id, week, period)
);
"""


def connect(db_path: Path) -> sqlite3.Connection:
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA)
    conn.commit()
