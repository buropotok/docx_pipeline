from __future__ import annotations

import sqlite3
from pathlib import Path

from docx_pipeline.config.settings import get_settings


def get_connection() -> sqlite3.Connection:
    settings = get_settings()
    settings.db_path.parent.mkdir(parents=True, exist_ok=True)
    # print(f"DB PATH = {settings.db_path}")
    conn = sqlite3.connect(settings.db_path)
    # print(conn.execute("PRAGMA database_list;").fetchall())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def ensure_db_exists() -> Path:
    settings = get_settings()
    settings.db_path.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(settings.db_path)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.close()

    return settings.db_path