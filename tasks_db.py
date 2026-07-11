"""
tasks_db.py — full CRUD, backed by SQLite.

Fully standalone module: no dependency on any other project (including the
HR Assistant app). Uses only Python's built-in sqlite3, so it needs no
server, no secrets/config file, and no extra setup beyond `pip install
streamlit`. The database is a single file (tasks.db) that lives next to
this module and persists across app restarts for as long as the app's
filesystem persists (true on a self-hosted/local deployment; some hosted
platforms reset the filesystem on redeploy, which would reset tasks.db too
— for guaranteed durability across redeploys, point DB_PATH at a
persistent volume instead).
"""

import os
import sqlite3
from datetime import datetime, date

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tasks.db")

VALID_STATUSES = ("Pending", "In Progress", "Completed")


def _get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Creates the tasks table if it doesn't exist yet. Safe to call on
    every page load."""
    conn = _get_conn()
    try:
        conn.execute(
            """CREATE TABLE IF NOT EXISTS tasks (
                task_id      INTEGER PRIMARY KEY AUTOINCREMENT,
                title        TEXT NOT NULL,
                description  TEXT,
                due_date     TEXT,
                status       TEXT NOT NULL DEFAULT 'Pending',
                created_at   TEXT NOT NULL DEFAULT (datetime('now')),
                completed_at TEXT
            )"""
        )
        conn.commit()
    finally:
        conn.close()


# --------------------------------------------------------------------------
# CRUD
# --------------------------------------------------------------------------

def create_task(title: str, description: str = "", due_date: date = None, status: str = "Pending") -> int:
    if not title or not title.strip():
        raise ValueError("Task title cannot be empty.")
    if status not in VALID_STATUSES:
        raise ValueError(f"Status must be one of {VALID_STATUSES}.")

    conn = _get_conn()
    try:
        cur = conn.execute(
            "INSERT INTO tasks (title, description, due_date, status) VALUES (?, ?, ?, ?)",
            (title.strip(), description or "", due_date.isoformat() if due_date else None, status),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def list_tasks(status_filter: str = None) -> list:
    """Returns all tasks (optionally filtered by status), ordered so that
    tasks with a due date come first (soonest first), then undated tasks."""
    conn = _get_conn()
    try:
        if status_filter and status_filter != "All":
            rows = conn.execute(
                "SELECT * FROM tasks WHERE status = ? "
                "ORDER BY (due_date IS NULL), due_date ASC, created_at ASC",
                (status_filter,),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM tasks ORDER BY (due_date IS NULL), due_date ASC, created_at ASC"
            ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_task(task_id: int) -> dict:
    conn = _get_conn()
    try:
        row = conn.execute("SELECT * FROM tasks WHERE task_id = ?", (task_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_task(
    task_id: int,
    title: str = None,
    description: str = None,
    due_date=None,
    status: str = None,
    _due_date_explicitly_cleared: bool = False,
) -> bool:
    """Updates only the fields explicitly provided (None = leave unchanged),
    except due_date: pass `_due_date_explicitly_cleared=True` to blank it out
    deliberately (since None is otherwise used to mean "don't touch this
    field"). Automatically stamps/clears completed_at when status changes
    to/from 'Completed'."""
    if status is not None and status not in VALID_STATUSES:
        raise ValueError(f"Status must be one of {VALID_STATUSES}.")

    fields, params = [], []
    if title is not None:
        if not title.strip():
            raise ValueError("Task title cannot be empty.")
        fields.append("title = ?")
        params.append(title.strip())
    if description is not None:
        fields.append("description = ?")
        params.append(description)
    if due_date is not None or _due_date_explicitly_cleared:
        fields.append("due_date = ?")
        params.append(due_date.isoformat() if due_date else None)
    if status is not None:
        fields.append("status = ?")
        params.append(status)
        fields.append("completed_at = ?")
        params.append(datetime.now().isoformat(timespec="seconds") if status == "Completed" else None)

    if not fields:
        return False

    params.append(task_id)
    conn = _get_conn()
    try:
        cur = conn.execute(f"UPDATE tasks SET {', '.join(fields)} WHERE task_id = ?", params)
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def delete_task(task_id: int) -> bool:
    conn = _get_conn()
    try:
        cur = conn.execute("DELETE FROM tasks WHERE task_id = ?", (task_id,))
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def delete_all_completed() -> int:
    """Convenience bulk-cleanup: removes every task already marked Completed."""
    conn = _get_conn()
    try:
        cur = conn.execute("DELETE FROM tasks WHERE status = 'Completed'")
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()
