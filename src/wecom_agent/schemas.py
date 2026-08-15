from __future__ import annotations

from dataclasses import asdict, dataclass, field
import hashlib
import json
from pathlib import Path
from typing import Any


class PlanValidationError(ValueError):
    """Raised when a model-produced send plan violates the input contract."""


@dataclass(frozen=True)
class RowSelection:
    rows: tuple[int, ...] = ()
    row_from: int | None = None
    row_to: int | None = None

    @classmethod
    def from_dict(cls, value: dict[str, Any] | None) -> "RowSelection":
        value = value or {}
        rows = tuple(int(row) for row in value.get("rows", []) or [])
        row_from = int(value["row_from"]) if value.get("row_from") is not None else None
        row_to = int(value["row_to"]) if value.get("row_to") is not None else None
        selection = cls(rows=rows, row_from=row_from, row_to=row_to)
        selection.validate()
        return selection

    def validate(self) -> None:
        if any(row < 2 for row in self.rows):
            raise PlanValidationError("Excel data rows must be 2 or greater")
        if len(set(self.rows)) != len(self.rows):
            raise PlanValidationError("Selected rows must not contain duplicates")
        if self.row_from is not None and self.row_from < 2:
            raise PlanValidationError("row_from must be 2 or greater")
        if self.row_to is not None and self.row_to < 2:
            raise PlanValidationError("row_to must be 2 or greater")
        if self.row_from is not None and self.row_to is not None and self.row_from > self.row_to:
            raise PlanValidationError("row_from cannot be greater than row_to")
        if self.rows and (self.row_from is not None or self.row_to is not None):
            raise PlanValidationError("Use either explicit rows or a row range, not both")


@dataclass(frozen=True)
class SendPlan:
    workbook: str | None = None
    lesson_workbook: str | None = None
    target_workbook: str | None = None
    lesson: str | None = None
    selection: RowSelection = field(default_factory=RowSelection)
    schedule_mode: str = "immediate"
    run_at: str | None = None
    resend: bool = False

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "SendPlan":
        allowed = {
            "workbook", "lesson_workbook", "target_workbook", "lesson", "selection",
            "schedule_mode", "run_at", "resend",
        }
        unknown = set(value) - allowed
        if unknown:
            raise PlanValidationError(f"Unsupported plan fields: {', '.join(sorted(unknown))}")
        plan = cls(
            workbook=_clean_optional(value.get("workbook")),
            lesson_workbook=_clean_optional(value.get("lesson_workbook")),
            target_workbook=_clean_optional(value.get("target_workbook")),
            lesson=_clean_optional(value.get("lesson")),
            selection=RowSelection.from_dict(value.get("selection")),
            schedule_mode=str(value.get("schedule_mode", "immediate")).strip().lower(),
            run_at=_clean_optional(value.get("run_at")),
            resend=bool(value.get("resend", False)),
        )
        plan.validate()
        return plan

    def validate(self) -> None:
        direct = bool(self.workbook)
        lesson_mode = bool(self.lesson_workbook or self.target_workbook or self.lesson)
        if direct == lesson_mode:
            raise PlanValidationError("Specify either workbook or the complete lesson workbook set")
        if lesson_mode and not (self.lesson_workbook and self.target_workbook and self.lesson):
            raise PlanValidationError("Lesson mode requires lesson_workbook, target_workbook, and lesson")
        if self.schedule_mode not in {"immediate", "workbook", "scheduled"}:
            raise PlanValidationError("schedule_mode must be immediate, workbook, or scheduled")
        if self.schedule_mode == "scheduled" and not self.run_at:
            raise PlanValidationError("scheduled mode requires run_at")
        if self.schedule_mode != "scheduled" and self.run_at:
            raise PlanValidationError("run_at is only valid in scheduled mode")

    def resolve_paths(self, workspace: Path) -> "SendPlan":
        workspace = workspace.resolve()

        def resolve(value: str | None) -> str | None:
            if not value:
                return None
            path = Path(value)
            path = path.resolve() if path.is_absolute() else (workspace / path).resolve()
            try:
                path.relative_to(workspace)
            except ValueError as exc:
                raise PlanValidationError(f"Workbook must stay inside workspace: {value}") from exc
            if not path.exists() or not path.is_file():
                raise PlanValidationError(f"Workbook does not exist: {value}")
            if path.suffix.lower() not in {".xlsx", ".xlsm"}:
                raise PlanValidationError(f"Unsupported workbook type: {value}")
            return str(path)

        return SendPlan(
            workbook=resolve(self.workbook),
            lesson_workbook=resolve(self.lesson_workbook),
            target_workbook=resolve(self.target_workbook),
            lesson=self.lesson,
            selection=self.selection,
            schedule_mode=self.schedule_mode,
            run_at=self.run_at,
            resend=self.resend,
        )

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)

    def canonical_json(self) -> str:
        return json.dumps(self.as_dict(), ensure_ascii=False, sort_keys=True, separators=(",", ":"))

    def plan_hash(self) -> str:
        return hashlib.sha256(self.canonical_json().encode("utf-8")).hexdigest()


def _clean_optional(value: Any) -> str | None:
    text = str(value).strip() if value is not None else ""
    return text or None
