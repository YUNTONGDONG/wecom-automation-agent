from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import subprocess
import sys
from typing import Any, Protocol

from .execution_policy import ApprovedTextTask


class ExecutionRunner(Protocol):
    def execute(self, task: ApprovedTextTask) -> dict[str, Any]: ...


@dataclass(frozen=True)
class FakeExecutionRunner:
    """Safe milestone runner: records intent but never imports or invokes GUI code."""

    def execute(self, task: ApprovedTextTask) -> dict[str, Any]:
        return {
            "mode": "supervised_fake",
            "status": "completed",
            "real_send": False,
            "row_number": task.row_number,
            "target": task.target,
            "message": "Fake Runner completed. WeCom was not opened and no message was sent.",
        }


class RealExecutionError(RuntimeError):
    """Raised when the supervised GUI sender cannot prove a successful send."""


@dataclass(frozen=True)
class RealWeComExecutionRunner:
    workspace: Path
    command_runner: Any = subprocess.run
    timeout_seconds: int = 300

    def execute(self, task: ApprovedTextTask) -> dict[str, Any]:
        sender = self.workspace / "scripts" / "send_from_excel_1v1_text.py"
        if not sender.is_file():
            raise RealExecutionError(f"WeCom sender not found: {sender}")
        command = [
            sys.executable,
            str(sender),
            "--folder", str(self.workspace),
            "--workbook", task.workbook,
            "--row", str(task.row_number),
            "--ignore-send-time",
            "--execute",
            "--yes",
            "--json",
            "--test-disclaimer", "",
            "--save-each-row",
            "--stop-on-error",
        ]
        completed = self.command_runner(
            command,
            cwd=self.workspace,
            text=True,
            capture_output=True,
            timeout=self.timeout_seconds,
            check=False,
        )
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout or "WeCom sender failed").strip()
            raise RealExecutionError(detail)
        payload = _last_json_object(completed.stdout)
        if int(payload.get("success", 0)) + int(payload.get("late_success", 0)) != 1:
            raise RealExecutionError(
                f"WeCom sender did not verify exactly one successful send: "
                f"success={payload.get('success', 0)}, failed={payload.get('failed', 0)}"
            )
        if int(payload.get("failed", 0)) != 0:
            raise RealExecutionError("WeCom sender reported a failed row")
        return {
            "mode": "supervised_real",
            "status": "completed",
            "real_send": True,
            "row_number": task.row_number,
            "target": task.target,
            "batch_id": payload.get("batch_id"),
            "run_dir": payload.get("run_dir"),
            "log_path": payload.get("log_path"),
            "evidence_manifest_path": payload.get("evidence_manifest_path"),
            "saved_path": payload.get("saved_path"),
        }


def _last_json_object(output: str) -> dict[str, Any]:
    decoder = json.JSONDecoder()
    for index in range(len(output) - 1, -1, -1):
        if output[index] != "{":
            continue
        try:
            value, end = decoder.raw_decode(output[index:])
        except json.JSONDecodeError:
            continue
        if output[index + end:].strip() == "" and isinstance(value, dict):
            return value
    raise RealExecutionError("WeCom sender did not return a JSON result")
