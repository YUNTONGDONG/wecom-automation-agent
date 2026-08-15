from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import secrets
from typing import Any


class ExecutionLocked(RuntimeError):
    """Raised when another desktop execution owns the workspace lock."""


@dataclass
class ExecutionLock:
    path: Path
    task_id: str
    _owner: str | None = None

    def acquire(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        owner = secrets.token_hex(16)
        payload = json.dumps({"task_id": self.task_id, "pid": os.getpid(), "owner": owner})
        try:
            descriptor = os.open(self.path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        except FileExistsError as exc:
            raise ExecutionLocked(f"Another execution owns lock: {self.path}") from exc
        try:
            os.write(descriptor, payload.encode("utf-8"))
        finally:
            os.close(descriptor)
        self._owner = owner

    def release(self) -> None:
        if not self._owner:
            return
        try:
            payload: dict[str, Any] = json.loads(self.path.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError):
            self._owner = None
            return
        if payload.get("owner") != self._owner:
            raise ExecutionLocked("Execution lock ownership changed; refusing to remove it")
        self.path.unlink()
        self._owner = None

    def __enter__(self) -> "ExecutionLock":
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        self.release()
