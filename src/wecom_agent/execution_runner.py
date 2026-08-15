from __future__ import annotations

from dataclasses import dataclass
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
