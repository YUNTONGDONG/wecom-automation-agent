from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import subprocess
import sys
from typing import Any, Callable

from .schemas import SendPlan


class ToolError(RuntimeError):
    """Raised when an Agent tool cannot safely complete."""


@dataclass(frozen=True)
class ToolDefinition:
    name: str
    description: str
    parameters: dict[str, Any]

    def as_openai_tool(self) -> dict[str, Any]:
        return {
            "type": "function",
            "name": self.name,
            "description": self.description,
            "parameters": self.parameters,
            "strict": True,
        }


PLAN_PARAMETERS: dict[str, Any] = {
    "type": "object",
    "properties": {
        "workbook": {"type": ["string", "null"]},
        "lesson_workbook": {"type": ["string", "null"]},
        "target_workbook": {"type": ["string", "null"]},
        "lesson": {"type": ["string", "null"]},
        "selection": {
            "type": "object",
            "properties": {
                "rows": {"type": "array", "items": {"type": "integer", "minimum": 2}},
                "row_from": {"type": ["integer", "null"], "minimum": 2},
                "row_to": {"type": ["integer", "null"], "minimum": 2},
            },
            "required": ["rows", "row_from", "row_to"],
            "additionalProperties": False,
        },
        "schedule_mode": {"type": "string", "enum": ["immediate", "workbook", "scheduled"]},
        "run_at": {"type": ["string", "null"]},
        "resend": {"type": "boolean"},
    },
    "required": [
        "workbook", "lesson_workbook", "target_workbook", "lesson", "selection",
        "schedule_mode", "run_at", "resend",
    ],
    "additionalProperties": False,
}


class WeComToolbox:
    def __init__(
        self,
        workspace: Path,
        command_runner: Callable[..., subprocess.CompletedProcess[str]] = subprocess.run,
    ) -> None:
        self.workspace = workspace.resolve()
        self.command_runner = command_runner
        self.sender = self.workspace / "scripts" / "send_from_excel_1v1_text.py"

    @property
    def definitions(self) -> list[ToolDefinition]:
        return [
            ToolDefinition(
                "preview_wecom_tasks",
                "Validate and preview an Excel-driven WeCom plan. This tool never sends messages.",
                PLAN_PARAMETERS,
            ),
        ]

    def dispatch(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        if name == "preview_wecom_tasks":
            return self.preview(arguments)
        raise ToolError(f"Unknown or forbidden tool: {name}")

    def preview(self, arguments: dict[str, Any]) -> dict[str, Any]:
        plan = SendPlan.from_dict(arguments).resolve_paths(self.workspace)
        command = self._command_for(plan)
        completed = self.command_runner(
            command,
            cwd=self.workspace,
            text=True,
            capture_output=True,
            timeout=60,
            check=False,
        )
        if completed.returncode != 0:
            raise ToolError((completed.stderr or completed.stdout or "Preview failed").strip())
        payload = _last_json_object(completed.stdout)
        return {
            "mode": "preview",
            "plan": plan.as_dict(),
            "plan_hash": plan.plan_hash(),
            "sendable_count": int(payload.get("sendable_rows", 0)),
            "skipped_count": int(payload.get("skipped_rows", 0)),
            "target_count": int(payload.get("target_count", 0)),
            "batch_id": payload.get("batch_id"),
            "run_dir": payload.get("run_dir"),
            "requires_confirmation": True,
        }

    def simulate_verified(self, plan: SendPlan) -> dict[str, Any]:
        return {
            "mode": "simulation",
            "status": "completed",
            "plan_hash": plan.plan_hash(),
            "message": "Simulation completed. WeCom was not opened and no message was sent.",
        }

    def _command_for(self, plan: SendPlan) -> list[str]:
        command = [sys.executable, str(self.sender), "--folder", str(self.workspace), "--json", "--test-disclaimer", ""]
        if plan.workbook:
            command.extend(["--workbook", plan.workbook])
        else:
            command.extend([
                "--lesson-workbook", plan.lesson_workbook or "",
                "--target-workbook", plan.target_workbook or "",
                "--lesson", plan.lesson or "",
            ])
        for row in plan.selection.rows:
            command.extend(["--row", str(row)])
        if plan.selection.row_from is not None:
            command.extend(["--row-from", str(plan.selection.row_from)])
        if plan.selection.row_to is not None:
            command.extend(["--row-to", str(plan.selection.row_to)])
        if plan.schedule_mode == "immediate":
            command.append("--ignore-send-time")
        elif plan.schedule_mode == "workbook":
            command.append("--respect-send-time")
        elif plan.run_at:
            command.extend(["--run-at", plan.run_at, "--respect-send-time"])
        if plan.resend:
            command.append("--resend")
        return command


def _last_json_object(output: str) -> dict[str, Any]:
    decoder = json.JSONDecoder()
    for index in range(len(output) - 1, -1, -1):
        if output[index] != "{":
            continue
        try:
            value, end = decoder.raw_decode(output[index:])
        except json.JSONDecodeError:
            continue
        if output[index + end :].strip() == "" and isinstance(value, dict):
            return value
    raise ToolError("Preview did not return a JSON object")
