from __future__ import annotations

import json
from typing import Any

from .model_client import ModelClient, ModelTurn
from .tools import ToolError, WeComToolbox


class AgentLoopError(RuntimeError):
    """Raised when the model exceeds the bounded orchestration loop."""


class WeComAgent:
    def __init__(self, model: ModelClient, toolbox: WeComToolbox, max_tool_rounds: int = 4) -> None:
        self.model = model
        self.toolbox = toolbox
        self.max_tool_rounds = max_tool_rounds

    def run(self, user_input: str) -> dict[str, Any]:
        tools = [definition.as_openai_tool() for definition in self.toolbox.definitions]
        turn = self.model.start(user_input, tools)
        events: list[dict[str, Any]] = []
        latest_result: dict[str, Any] | None = None

        for _round in range(self.max_tool_rounds + 1):
            if not turn.tool_calls:
                return {"message": turn.text, "tool_events": events, "latest_result": latest_result}
            outputs: list[dict[str, Any]] = []
            for call in turn.tool_calls:
                try:
                    result = self.toolbox.dispatch(call.name, call.arguments)
                    latest_result = result
                    event = {"tool": call.name, "status": "ok", "result": result}
                    output = result
                except Exception as exc:
                    event = {"tool": call.name, "status": "error", "error": str(exc)}
                    output = {"error": str(exc)}
                events.append(event)
                outputs.append({
                    "type": "function_call_output",
                    "call_id": call.call_id,
                    "output": json.dumps(output, ensure_ascii=False),
                })
            turn = self.model.continue_with(outputs, tools)
        raise AgentLoopError(f"Model exceeded {self.max_tool_rounds} tool rounds")
