from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sqlite3
from typing import Any, Iterator
import uuid

from .permissions import Approval, PermissionDenied
from .state_store import TaskState, VALID_TRANSITIONS


class DuplicateExecution(RuntimeError):
    """Raised when an idempotency key has already started execution."""


class SQLiteStateStore:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    @contextmanager
    def connection(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self.path, timeout=10, isolation_level=None)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        try:
            yield connection
        finally:
            connection.close()

    def create_task(self, state: TaskState) -> None:
        now = _now()
        with self.connection() as connection:
            connection.execute(
                """
                INSERT INTO tasks (
                    task_id, status, plan_json, plan_hash, preview_json, snapshot_json,
                    result_json, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    state.task_id,
                    state.status,
                    _dump(state.plan),
                    state.plan_hash,
                    _dump(state.preview),
                    _dump(state.snapshot),
                    _dump(state.result),
                    now,
                    now,
                ),
            )
            self._audit(connection, state.task_id, "task_created", {"status": state.status})

    def load_task(self, task_id: str) -> TaskState:
        with self.connection() as connection:
            row = connection.execute("SELECT * FROM tasks WHERE task_id = ?", (task_id,)).fetchone()
        if row is None:
            raise KeyError(f"Unknown task: {task_id}")
        return TaskState(
            task_id=row["task_id"],
            status=row["status"],
            plan=_load(row["plan_json"]),
            plan_hash=row["plan_hash"],
            preview=_load(row["preview_json"]),
            snapshot=_load(row["snapshot_json"]),
            result=_load(row["result_json"]),
        )

    def transition(self, task_id: str, new_status: str, result: dict[str, Any] | None = None) -> TaskState:
        with self.connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute("SELECT status FROM tasks WHERE task_id = ?", (task_id,)).fetchone()
            if row is None:
                connection.execute("ROLLBACK")
                raise KeyError(f"Unknown task: {task_id}")
            current = row["status"]
            if new_status not in VALID_TRANSITIONS.get(current, set()):
                connection.execute("ROLLBACK")
                raise ValueError(f"Invalid task transition: {current} -> {new_status}")
            connection.execute(
                "UPDATE tasks SET status = ?, result_json = COALESCE(?, result_json), updated_at = ? WHERE task_id = ?",
                (new_status, _dump(result), _now(), task_id),
            )
            self._audit(connection, task_id, "status_changed", {"from": current, "to": new_status})
            connection.execute("COMMIT")
        return self.load_task(task_id)

    def save_approval(self, approval: Approval) -> None:
        with self.connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            connection.execute(
                """
                INSERT INTO approvals (
                    approval_id, task_id, plan_hash, token_hash, issued_at, expires_at, consumed_at
                ) VALUES (?, ?, ?, ?, ?, ?, NULL)
                """,
                (
                    uuid.uuid4().hex,
                    approval.task_id,
                    approval.plan_hash,
                    approval.token_hash,
                    approval.issued_at,
                    approval.expires_at,
                ),
            )
            self._audit(connection, approval.task_id, "approval_issued", {"expires_at": approval.expires_at})
            connection.execute("COMMIT")

    def consume_approval(self, approval: Approval, *, now: datetime | None = None) -> None:
        token_hash = hashlib.sha256(approval.token.encode("utf-8")).hexdigest()
        now_text = (now or datetime.now(timezone.utc)).astimezone(timezone.utc).isoformat()
        with self.connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                """
                SELECT approval_id, expires_at, consumed_at FROM approvals
                WHERE task_id = ? AND plan_hash = ? AND token_hash = ?
                """,
                (approval.task_id, approval.plan_hash, token_hash),
            ).fetchone()
            if row is None:
                connection.execute("ROLLBACK")
                raise PermissionDenied("Approval is not registered")
            if row["consumed_at"] is not None:
                connection.execute("ROLLBACK")
                raise PermissionDenied("Approval has already been consumed")
            if now_text >= row["expires_at"]:
                connection.execute("ROLLBACK")
                raise PermissionDenied("Approval has expired")
            connection.execute(
                "UPDATE approvals SET consumed_at = ? WHERE approval_id = ? AND consumed_at IS NULL",
                (now_text, row["approval_id"]),
            )
            self._audit(connection, approval.task_id, "approval_consumed", {})
            connection.execute("COMMIT")

    def start_execution(self, task_id: str, idempotency_key: str, mode: str) -> str:
        execution_id = uuid.uuid4().hex
        with self.connection() as connection:
            try:
                connection.execute(
                    """
                    INSERT INTO executions (
                        execution_id, task_id, idempotency_key, mode, status, started_at
                    ) VALUES (?, ?, ?, ?, 'running', ?)
                    """,
                    (execution_id, task_id, idempotency_key, mode, _now()),
                )
            except sqlite3.IntegrityError as exc:
                raise DuplicateExecution("This exact task snapshot has already started") from exc
            self._audit(connection, task_id, "execution_started", {"execution_id": execution_id, "mode": mode})
        return execution_id

    def begin_execution_with_approval(
        self,
        approval: Approval,
        idempotency_key: str,
        mode: str,
        *,
        now: datetime | None = None,
    ) -> str:
        """Atomically consume approval, reserve idempotency, and enter execution state."""
        execution_id = uuid.uuid4().hex
        token_hash = hashlib.sha256(approval.token.encode("utf-8")).hexdigest()
        now_text = (now or datetime.now(timezone.utc)).astimezone(timezone.utc).isoformat()
        next_status = "SIMULATING" if mode == "simulation" else "EXECUTING"
        with self.connection() as connection:
            connection.execute("BEGIN IMMEDIATE")
            task = connection.execute(
                "SELECT status, plan_hash FROM tasks WHERE task_id = ?",
                (approval.task_id,),
            ).fetchone()
            if task is None:
                connection.execute("ROLLBACK")
                raise KeyError(f"Unknown task: {approval.task_id}")
            if task["status"] != "AUTHORIZED" or task["plan_hash"] != approval.plan_hash:
                connection.execute("ROLLBACK")
                raise PermissionDenied("Task is not authorized for this exact plan")
            stored = connection.execute(
                """
                SELECT approval_id, expires_at, consumed_at FROM approvals
                WHERE task_id = ? AND plan_hash = ? AND token_hash = ?
                """,
                (approval.task_id, approval.plan_hash, token_hash),
            ).fetchone()
            if stored is None or stored["consumed_at"] is not None:
                connection.execute("ROLLBACK")
                raise PermissionDenied("Approval is missing or already consumed")
            if now_text >= stored["expires_at"]:
                connection.execute("ROLLBACK")
                raise PermissionDenied("Approval has expired")
            try:
                connection.execute(
                    """
                    INSERT INTO executions (
                        execution_id, task_id, idempotency_key, mode, status, started_at
                    ) VALUES (?, ?, ?, ?, 'running', ?)
                    """,
                    (execution_id, approval.task_id, idempotency_key, mode, now_text),
                )
            except sqlite3.IntegrityError as exc:
                connection.execute("ROLLBACK")
                raise DuplicateExecution("This exact task snapshot has already started") from exc
            connection.execute(
                "UPDATE approvals SET consumed_at = ? WHERE approval_id = ? AND consumed_at IS NULL",
                (now_text, stored["approval_id"]),
            )
            connection.execute(
                "UPDATE tasks SET status = ?, updated_at = ? WHERE task_id = ?",
                (next_status, now_text, approval.task_id),
            )
            self._audit(connection, approval.task_id, "approval_consumed", {})
            self._audit(connection, approval.task_id, "execution_started", {"execution_id": execution_id, "mode": mode})
            self._audit(connection, approval.task_id, "status_changed", {"from": "AUTHORIZED", "to": next_status})
            connection.execute("COMMIT")
        return execution_id

    def complete_execution(self, execution_id: str, status: str, result: dict[str, Any]) -> None:
        with self.connection() as connection:
            row = connection.execute("SELECT task_id FROM executions WHERE execution_id = ?", (execution_id,)).fetchone()
            if row is None:
                raise KeyError(f"Unknown execution: {execution_id}")
            connection.execute(
                "UPDATE executions SET status = ?, completed_at = ?, result_json = ? WHERE execution_id = ?",
                (status, _now(), _dump(result), execution_id),
            )
            self._audit(connection, row["task_id"], "execution_completed", {"execution_id": execution_id, "status": status})

    def reserve_delivery(
        self,
        delivery_key: str,
        execution_id: str,
        row_number: int,
        target: str,
        message_hash: str,
    ) -> None:
        with self.connection() as connection:
            try:
                connection.execute(
                    """
                    INSERT INTO deliveries (
                        delivery_key, execution_id, row_number, target, message_hash, status, created_at
                    ) VALUES (?, ?, ?, ?, ?, 'reserved', ?)
                    """,
                    (delivery_key, execution_id, row_number, target, message_hash, _now()),
                )
            except sqlite3.IntegrityError as exc:
                raise DuplicateExecution("Delivery has already been reserved or attempted") from exc

    def audit_events(self, task_id: str) -> list[dict[str, Any]]:
        with self.connection() as connection:
            rows = connection.execute(
                "SELECT event_type, event_json, created_at FROM audit_events WHERE task_id = ? ORDER BY event_id",
                (task_id,),
            ).fetchall()
        return [{"type": row["event_type"], "data": _load(row["event_json"]), "created_at": row["created_at"]} for row in rows]

    def _initialize(self) -> None:
        with self.connection() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS tasks (
                    task_id TEXT PRIMARY KEY,
                    status TEXT NOT NULL,
                    plan_json TEXT,
                    plan_hash TEXT,
                    preview_json TEXT,
                    snapshot_json TEXT,
                    result_json TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS approvals (
                    approval_id TEXT PRIMARY KEY,
                    task_id TEXT NOT NULL REFERENCES tasks(task_id),
                    plan_hash TEXT NOT NULL,
                    token_hash TEXT NOT NULL UNIQUE,
                    issued_at TEXT NOT NULL,
                    expires_at TEXT NOT NULL,
                    consumed_at TEXT
                );
                CREATE TABLE IF NOT EXISTS executions (
                    execution_id TEXT PRIMARY KEY,
                    task_id TEXT NOT NULL REFERENCES tasks(task_id),
                    idempotency_key TEXT NOT NULL UNIQUE,
                    mode TEXT NOT NULL,
                    status TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    completed_at TEXT,
                    result_json TEXT
                );
                CREATE TABLE IF NOT EXISTS deliveries (
                    delivery_key TEXT PRIMARY KEY,
                    execution_id TEXT NOT NULL REFERENCES executions(execution_id),
                    row_number INTEGER NOT NULL,
                    target TEXT NOT NULL,
                    message_hash TEXT NOT NULL,
                    status TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS audit_events (
                    event_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL REFERENCES tasks(task_id),
                    event_type TEXT NOT NULL,
                    event_json TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );
                """
            )

    def _audit(self, connection: sqlite3.Connection, task_id: str, event_type: str, data: dict[str, Any]) -> None:
        connection.execute(
            "INSERT INTO audit_events (task_id, event_type, event_json, created_at) VALUES (?, ?, ?, ?)",
            (task_id, event_type, _dump(data) or "{}", _now()),
        )


def _dump(value: Any) -> str | None:
    return None if value is None else json.dumps(value, ensure_ascii=False, sort_keys=True)


def _load(value: str | None) -> Any:
    return None if value is None else json.loads(value)


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()
