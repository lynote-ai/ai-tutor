"""SQLite storage for the shelf and generated courses."""

from __future__ import annotations

import json
import sqlite3
import threading
from pathlib import Path

_lock = threading.Lock()
_conn: sqlite3.Connection | None = None


def connect(path: Path) -> None:
    global _conn
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS shelf_item (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            kind TEXT NOT NULL,
            text TEXT NOT NULL,
            spans_json TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS course (
            id TEXT PRIMARY KEY,
            title TEXT NOT NULL,
            goal TEXT NOT NULL,
            depth TEXT NOT NULL,
            status TEXT NOT NULL,
            stage TEXT NOT NULL,
            message TEXT NOT NULL,
            progress INTEGER NOT NULL,
            warning TEXT NOT NULL,
            overview TEXT NOT NULL,
            sample INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS course_material (
            id TEXT PRIMARY KEY,
            course_id TEXT NOT NULL,
            name TEXT NOT NULL,
            kind TEXT NOT NULL,
            text TEXT NOT NULL,
            spans_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS chapter (
            id TEXT PRIMARY KEY,
            course_id TEXT NOT NULL,
            position INTEGER NOT NULL,
            title TEXT NOT NULL,
            overview TEXT NOT NULL,
            lesson_md TEXT NOT NULL,
            summary_md TEXT NOT NULL,
            citations_json TEXT NOT NULL,
            questions_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS chapter_state (
            course_id TEXT NOT NULL,
            chapter_id TEXT NOT NULL,
            opened INTEGER NOT NULL DEFAULT 0,
            quiz_choice TEXT NOT NULL DEFAULT '',
            quiz_correct INTEGER NOT NULL DEFAULT 0,
            note TEXT NOT NULL DEFAULT '',
            PRIMARY KEY (course_id, chapter_id)
        );
        """
    )
    conn.commit()
    _conn = conn


def _db() -> sqlite3.Connection:
    if _conn is None:
        raise RuntimeError("Database is not connected.")
    return _conn


def execute(sql: str, params: tuple = ()) -> None:
    with _lock:
        _db().execute(sql, params)
        _db().commit()



def execute_batch(statements: list[tuple[str, tuple]]) -> None:
    """Execute multiple SQL statements in one transaction with rollback on error."""
    with _lock:
        try:
            for sql, params in statements:
                _db().execute(sql, params)
            _db().commit()
        except Exception:
            _db().rollback()
            raise


def query_all(sql: str, params: tuple = ()) -> list[dict]:
    with _lock:
        rows = _db().execute(sql, params).fetchall()
    return [dict(row) for row in rows]


def query_one(sql: str, params: tuple = ()) -> dict | None:
    rows = query_all(sql, params)
    return rows[0] if rows else None


def dump_json(value) -> str:
    return json.dumps(value, ensure_ascii=False)


def load_json(raw: str, fallback):
    try:
        return json.loads(raw or "")
    except json.JSONDecodeError:
        return fallback
