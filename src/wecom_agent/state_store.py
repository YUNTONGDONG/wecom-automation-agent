from __future__ import annotations

from dataclasses import dataclass
from typing import Any


VALID_TRANSITIONS = {
    "DRAFT": {"PREVIEWED"},
    "PREVIEWED": {"AWAITING_CONFIRMATION"},
    "AWAITING_CONFIRMATION": {"AUTHORIZED"},
    "AUTHORIZED": {"SIMULATING", "EXECUTING"},
    "SIMULATING": {"COMPLETED", "FAILED"},
    "EXECUTING": {"COMPLETED", "PARTIAL", "MANUAL_REVIEW", "FAILED"},
    "COMPLETED": set(),
    "PARTIAL": set(),
    "MANUAL_REVIEW": set(),
    "FAILED": set(),
}


@dataclass
class TaskState:
    task_id: str
    status: str = "DRAFT"
    plan: dict[str, Any] | None = None
    plan_hash: str | None = None
    preview: dict[str, Any] | None = None
    snapshot: dict[str, Any] | None = None
    result: dict[str, Any] | None = None

    def transition(self, status: str) -> None:
        if status not in VALID_TRANSITIONS.get(self.status, set()):
            raise ValueError(f"Invalid task transition: {self.status} -> {status}")
        self.status = status
