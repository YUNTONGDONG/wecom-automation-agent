from __future__ import annotations

from dataclasses import dataclass
import json
import os
import re
from typing import Any, Protocol


@dataclass(frozen=True)
class ToolCall:
    call_id: str
    name: str
    arguments: dict[str, Any]


@dataclass(frozen=True)
class ModelTurn:
    text: str = ""
    tool_calls: tuple[ToolCall, ...] = ()
    response_id: str | None = None


class ModelClient(Protocol):
    def start(self, user_input: str, tools: list[dict[str, Any]]) -> ModelTurn: ...

    def continue_with(self, outputs: list[dict[str, Any]], tools: list[dict[str, Any]]) -> ModelTurn: ...


SYSTEM_PROMPT = """You are a safety-focused WeCom planning agent. Real sending is unavailable.
Use at most one preview_wecom_tasks call per turn; it can validate one unambiguous plan and never sends messages.
If no scheduling instruction is given, use schedule_mode=immediate. Use workbook only when the user explicitly asks to respect times stored in the workbook.
If the request names multiple workbooks without all lesson-mode roles, do not call a tool; ask the user to identify the intended workbook or roles.
Preserve user-provided row values and order exactly. Never reorder, repair, or infer an invalid row range; ask for clarification instead.
Never invent workbook paths, recipients, lesson numbers, rows, approval tokens, or missing plan values.
Never claim a message was sent. Approval and execution are handled outside the model by a human-operated CLI.
After a tool result, return a concise summary that preserves any error or confirmation requirement.
"""


class OpenAIResponsesClient:
    """Thin Responses API adapter. No request is made until start() is called."""

    def __init__(self, model: str | None = None, max_retries: int = 2) -> None:
        try:
            from openai import OpenAI
        except ImportError as exc:
            raise RuntimeError("Install the optional 'openai' dependency to use the live model client") from exc
        self.client = OpenAI(max_retries=max_retries)
        self.model = model or os.environ.get("AGENT_MODEL", "gpt-5.6-terra")
        self.previous_response_id: str | None = None

    def start(self, user_input: str, tools: list[dict[str, Any]]) -> ModelTurn:
        response = self.client.responses.create(
            model=self.model,
            instructions=SYSTEM_PROMPT,
            input=user_input,
            tools=tools,
            tool_choice="auto",
            reasoning={"effort": "low"},
        )
        self.previous_response_id = response.id
        return _parse_response(response)

    def continue_with(self, outputs: list[dict[str, Any]], tools: list[dict[str, Any]]) -> ModelTurn:
        if not self.previous_response_id:
            raise RuntimeError("Cannot continue before starting a response")
        response = self.client.responses.create(
            model=self.model,
            previous_response_id=self.previous_response_id,
            input=outputs,
            tools=tools,
            tool_choice="auto",
            reasoning={"effort": "low"},
        )
        self.previous_response_id = response.id
        return _parse_response(response)


class RuleBasedMockClient:
    """Deterministic local model double for demos and tests."""

    def __init__(self, default_workbook: str = "examples/example_tasks.xlsx") -> None:
        self.default_workbook = default_workbook
        self.started = False

    def start(self, user_input: str, tools: list[dict[str, Any]]) -> ModelTurn:
        self.started = True
        normalized = user_input.lower()
        forbidden = (
            "跳过验证", "绕过审批", "直接发送", "真实发送", "执行shell", "执行 shell",
            "运行命令", "approval token", "ignore safety", "ignore previous",
        )
        if any(phrase in normalized for phrase in forbidden):
            return ModelTurn(text="该请求试图绕过安全边界；我只能生成并验证预览计划。")

        row_range = re.search(r"(?:第)?(\d+)\s*(?:到|至|-)\s*(?:第)?(\d+)\s*行", user_input)
        explicit_rows = tuple(int(value) for value in re.findall(r"第(\d+)行", user_input)) if not row_range else ()
        workbook_matches = re.findall(r"([^\s，,]+\.(?:xlsx|xlsm))", user_input, flags=re.IGNORECASE)
        if len(workbook_matches) > 1:
            return ModelTurn(text="检测到多个工作簿，请明确使用单工作簿模式，或分别提供课程表、目标表和课程名称。")
        workbook = workbook_matches[0] if workbook_matches else None
        if not workbook and ("示例" in user_input or "example" in normalized):
            workbook = self.default_workbook
        if not workbook:
            return ModelTurn(text="请提供工作区内的 Excel 工作簿路径后再预览。")

        scheduled_at = re.search(r"(20\d{2}-\d{2}-\d{2})[ T](\d{2}:\d{2}(?::\d{2})?)", user_input)
        if scheduled_at:
            schedule_mode = "scheduled"
            run_at = f"{scheduled_at.group(1)} {scheduled_at.group(2)}"
        elif "按表" in user_input or "表内时间" in user_input or "指定时间" in user_input:
            schedule_mode = "workbook"
            run_at = None
        else:
            schedule_mode = "immediate"
            run_at = None
        arguments = {
            "workbook": workbook,
            "lesson_workbook": None,
            "target_workbook": None,
            "lesson": None,
            "selection": {
                "rows": list(explicit_rows),
                "row_from": int(row_range.group(1)) if row_range else None,
                "row_to": int(row_range.group(2)) if row_range else None,
            },
            "schedule_mode": schedule_mode,
            "run_at": run_at,
            "resend": "重发" in user_input,
        }
        return ModelTurn(tool_calls=(ToolCall("mock-preview-1", "preview_wecom_tasks", arguments),))

    def continue_with(self, outputs: list[dict[str, Any]], tools: list[dict[str, Any]]) -> ModelTurn:
        payload = json.loads(outputs[-1]["output"])
        if payload.get("error"):
            return ModelTurn(text=f"预览失败：{payload['error']}")
        return ModelTurn(
            text=(
                f"预览完成：可发送 {payload['sendable_count']} 行，跳过 {payload['skipped_count']} 行，"
                f"目标 {payload['target_count']} 个。当前未操作企业微信，需要确认后才能模拟执行。"
            )
        )


def _parse_response(response: Any) -> ModelTurn:
    calls: list[ToolCall] = []
    for item in response.output:
        if getattr(item, "type", "") != "function_call":
            continue
        try:
            arguments = json.loads(item.arguments)
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"Model returned invalid tool arguments for {item.name}") from exc
        calls.append(ToolCall(item.call_id, item.name, arguments))
    return ModelTurn(
        text=getattr(response, "output_text", "") or "",
        tool_calls=tuple(calls),
        response_id=getattr(response, "id", None),
    )
