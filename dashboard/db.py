from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Iterator


DEFAULT_DB_PATH = Path("data") / "inverter_dashboard.db"


SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS imports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_hash TEXT NOT NULL UNIQUE,
    filename TEXT NOT NULL,
    week_start TEXT,
    week_end TEXT,
    imported_at TEXT NOT NULL,
    primary_sheet TEXT,
    source_unit TEXT,
    unit_multiplier REAL NOT NULL DEFAULT 1,
    status TEXT NOT NULL DEFAULT 'completed',
    workbook_meta_json TEXT NOT NULL DEFAULT '{}'
);

CREATE TABLE IF NOT EXISTS sheets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    state TEXT NOT NULL,
    max_row INTEGER NOT NULL,
    max_column INTEGER NOT NULL,
    header_row INTEGER,
    UNIQUE(import_id, name)
);

CREATE TABLE IF NOT EXISTS raw_cells (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    sheet_name TEXT NOT NULL,
    row_index INTEGER NOT NULL,
    column_index INTEGER NOT NULL,
    coordinate TEXT NOT NULL,
    raw_value TEXT,
    cached_value TEXT,
    data_type TEXT,
    number_format TEXT,
    UNIQUE(import_id, sheet_name, coordinate)
);

CREATE TABLE IF NOT EXISTS manager_plans (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    manager TEXT NOT NULL,
    scope TEXT NOT NULL,
    year INTEGER NOT NULL,
    quarter INTEGER,
    metric TEXT NOT NULL,
    amount REAL NOT NULL,
    source_sheet TEXT NOT NULL,
    source_row INTEGER NOT NULL,
    source_cell TEXT NOT NULL,
    UNIQUE(import_id, manager, scope, year, quarter, metric, source_cell)
);

CREATE TABLE IF NOT EXISTS project_snapshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    business_key TEXT NOT NULL,
    source_sheet TEXT NOT NULL,
    source_row INTEGER NOT NULL,
    manager TEXT,
    customer TEXT,
    final_customer TEXT,
    project_name TEXT NOT NULL,
    channel TEXT,
    industry TEXT,
    equipment_type TEXT,
    project_amount REAL,
    stage TEXT,
    contract_date TEXT,
    planned_ship_date TEXT,
    actual_ship_date TEXT,
    planned_year INTEGER,
    planned_quarter INTEGER,
    probability REAL,
    risk_financial INTEGER NOT NULL DEFAULT 0,
    risk_technical INTEGER NOT NULL DEFAULT 0,
    risk_production INTEGER NOT NULL DEFAULT 0,
    risk_procurement INTEGER NOT NULL DEFAULT 0,
    risk_schedule INTEGER NOT NULL DEFAULT 0,
    next_owner TEXT,
    next_action TEXT,
    next_action_due TEXT,
    comment TEXT,
    current_category TEXT,
    source_text TEXT NOT NULL,
    source_values_json TEXT NOT NULL DEFAULT '{}',
    last_updated TEXT,
    duplicate_group TEXT,
    is_ambiguous INTEGER NOT NULL DEFAULT 0,
    UNIQUE(import_id, source_sheet, source_row)
);

CREATE INDEX IF NOT EXISTS idx_project_snapshots_import
    ON project_snapshots(import_id);
CREATE INDEX IF NOT EXISTS idx_project_snapshots_key
    ON project_snapshots(import_id, business_key);
CREATE INDEX IF NOT EXISTS idx_project_snapshots_manager
    ON project_snapshots(import_id, manager);

CREATE TABLE IF NOT EXISTS period_values (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    business_key TEXT NOT NULL,
    source_sheet TEXT NOT NULL,
    source_row INTEGER NOT NULL,
    year INTEGER NOT NULL,
    quarter INTEGER NOT NULL,
    metric TEXT NOT NULL,
    amount REAL NOT NULL,
    source_cell TEXT NOT NULL,
    raw_amount REAL,
    UNIQUE(import_id, source_sheet, source_row, source_cell)
);

CREATE INDEX IF NOT EXISTS idx_period_values_import
    ON period_values(import_id, year, quarter, metric);
CREATE INDEX IF NOT EXISTS idx_period_values_key
    ON period_values(import_id, business_key);

CREATE TABLE IF NOT EXISTS project_overrides (
    business_key TEXT PRIMARY KEY,
    manager TEXT,
    customer TEXT,
    final_customer TEXT,
    project_name TEXT,
    channel TEXT,
    industry TEXT,
    equipment_type TEXT,
    stage TEXT,
    contract_date TEXT,
    planned_ship_date TEXT,
    actual_ship_date TEXT,
    probability REAL,
    risk_financial INTEGER,
    risk_technical INTEGER,
    risk_production INTEGER,
    risk_procurement INTEGER,
    risk_schedule INTEGER,
    next_owner TEXT,
    next_action TEXT,
    next_action_due TEXT,
    comment TEXT,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS review_queue (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    business_key TEXT,
    issue_type TEXT NOT NULL,
    details TEXT NOT NULL,
    resolved INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    UNIQUE(import_id, business_key, issue_type, details)
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
"""


def connect(path: str | Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    db_path = Path(path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA journal_mode = WAL")
    return connection


def initialize(connection: sqlite3.Connection) -> None:
    connection.executescript(SCHEMA)
    connection.commit()


@contextmanager
def transaction(connection: sqlite3.Connection) -> Iterator[sqlite3.Connection]:
    try:
        connection.execute("BEGIN")
        yield connection
    except Exception:
        connection.rollback()
        raise
    else:
        connection.commit()


def query_all(
    connection: sqlite3.Connection, sql: str, params: Iterable[Any] = ()
) -> list[sqlite3.Row]:
    return list(connection.execute(sql, tuple(params)).fetchall())


def latest_import_id(connection: sqlite3.Connection) -> int | None:
    row = connection.execute(
        "SELECT id FROM imports WHERE status = 'completed' "
        "ORDER BY COALESCE(week_end, imported_at) DESC, id DESC LIMIT 1"
    ).fetchone()
    return int(row["id"]) if row else None


def list_imports(connection: sqlite3.Connection) -> list[dict[str, Any]]:
    return [
        dict(row)
        for row in query_all(
            connection,
            "SELECT id, filename, week_start, week_end, imported_at, primary_sheet, "
            "source_unit, unit_multiplier FROM imports WHERE status = 'completed' "
            "ORDER BY COALESCE(week_end, imported_at) DESC, id DESC",
        )
    ]


OVERRIDE_FIELDS = {
    "manager",
    "customer",
    "final_customer",
    "project_name",
    "channel",
    "industry",
    "equipment_type",
    "stage",
    "contract_date",
    "planned_ship_date",
    "actual_ship_date",
    "probability",
    "risk_financial",
    "risk_technical",
    "risk_production",
    "risk_procurement",
    "risk_schedule",
    "next_owner",
    "next_action",
    "next_action_due",
    "comment",
}


def save_override(
    connection: sqlite3.Connection, business_key: str, values: dict[str, Any]
) -> None:
    cleaned = {key: value for key, value in values.items() if key in OVERRIDE_FIELDS}
    if not cleaned:
        return
    now = datetime.now().isoformat(timespec="seconds")
    existing = connection.execute(
        "SELECT * FROM project_overrides WHERE business_key = ?", (business_key,)
    ).fetchone()
    record = {field: None for field in OVERRIDE_FIELDS}
    if existing:
        record.update({field: existing[field] for field in OVERRIDE_FIELDS})
    record.update(cleaned)
    record["updated_at"] = now
    fields = ["business_key", *sorted(OVERRIDE_FIELDS), "updated_at"]
    placeholders = ", ".join("?" for _ in fields)
    updates = ", ".join(
        f"{field}=excluded.{field}" for field in fields if field != "business_key"
    )
    connection.execute(
        f"INSERT INTO project_overrides ({', '.join(fields)}) VALUES ({placeholders}) "
        f"ON CONFLICT(business_key) DO UPDATE SET {updates}",
        [business_key, *[record[field] for field in sorted(OVERRIDE_FIELDS)], now],
    )
    connection.commit()


def set_setting(connection: sqlite3.Connection, key: str, value: Any) -> None:
    encoded = json.dumps(value, ensure_ascii=False)
    connection.execute(
        "INSERT INTO settings(key, value, updated_at) VALUES(?, ?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at",
        (key, encoded, datetime.now().isoformat(timespec="seconds")),
    )
    connection.commit()


def get_setting(
    connection: sqlite3.Connection, key: str, default: Any = None
) -> Any:
    row = connection.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    if not row:
        return default
    try:
        return json.loads(row["value"])
    except json.JSONDecodeError:
        return default
