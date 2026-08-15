from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import secrets
import uuid

from .agent import WeComAgent
from .database import SQLiteStateStore
from .execution_lock import ExecutionLock
from .execution_policy import ApprovedTextTask, SupervisedExecutionPolicy, parse_allowed_targets
from .execution_runner import RealWeComExecutionRunner
from .model_client import OpenAIResponsesClient, RuleBasedMockClient
from .permissions import ApprovalManager
from .schemas import SendPlan
from .snapshots import (
    PlanSnapshot,
    assert_snapshot_unchanged,
    capture_plan_snapshot,
    execution_idempotency_key,
)
from .state_store import TaskState
from .tools import WeComToolbox


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Safety-gated WeCom Agent MVP")
    parser.add_argument("--workspace", default=".")
    subparsers = parser.add_subparsers(dest="command", required=True)

    preview = subparsers.add_parser("preview", help="Interpret a request and run a safe preview")
    preview.add_argument("request")
    preview.add_argument("--live", action="store_true", help="Use the OpenAI Responses API instead of the local mock")
    preview.add_argument("--model", default=None)

    send = subparsers.add_parser("send", help="Preview, confirm, approve, and send one safe workbook in one flow")
    send.add_argument("workbook")
    send.add_argument("--yes", action="store_true", help="Skip the interactive y/N prompt")
    send.add_argument("--simulate", action="store_true", help="Complete the one-click flow without opening WeCom")

    allow = subparsers.add_parser("allow", help="Add an exact WeCom contact name to the local allowlist")
    allow.add_argument("target")

    approve = subparsers.add_parser("approve", help="Approve the exact previewed plan for ten minutes")
    approve.add_argument("task_id")

    status = subparsers.add_parser("status", help="Inspect persisted task state and audit events")
    status.add_argument("task_id")

    simulate = subparsers.add_parser("simulate", help="Consume approval and simulate once; never opens WeCom")
    simulate.add_argument("task_id")
    simulate.add_argument("--approval-token", required=True)
    execute = subparsers.add_parser("execute", help="Send one approved, allowlisted plain-text task through WeCom")
    execute.add_argument("task_id")
    execute.add_argument("--approval-token", required=True)
    return parser


def main() -> None:
    args = build_parser().parse_args()
    workspace = Path(args.workspace).resolve()
    state_dir = workspace / ".agent-state"
    database = SQLiteStateStore(state_dir / "agent.db")
    approvals = ApprovalManager(_approval_secret(state_dir))
    toolbox = WeComToolbox(workspace)

    if args.command == "allow":
        targets = set(_allowed_targets(state_dir))
        target = args.target.strip()
        if not target:
            raise SystemExit("Target name must not be empty")
        targets.add(target)
        path = state_dir / "allowed-targets.txt"
        path.write_text("\n".join(sorted(targets)) + "\n", encoding="utf-8")
        try:
            path.chmod(0o600)
        except OSError:
            pass
        print(json.dumps({"status": "configured", "target": target, "allowlist_path": str(path)}, ensure_ascii=False, indent=2))
        return

    if args.command == "send":
        plan = SendPlan.from_dict({"workbook": args.workbook}).resolve_paths(workspace)
        preview = toolbox.preview(plan.as_dict())
        snapshot = capture_plan_snapshot(plan)
        task_id = f"task-{uuid.uuid4().hex[:10]}"
        state = TaskState(
            task_id=task_id,
            status="AWAITING_CONFIRMATION",
            plan=plan.as_dict(),
            plan_hash=plan.plan_hash(),
            preview=preview,
            snapshot=snapshot.as_dict(),
        )
        database.create_task(state)
        policy = SupervisedExecutionPolicy(_allowed_targets(state_dir))
        approved_text_task = policy.validate(plan)
        print("\n即将处理企业微信任务")
        print(f"  联系人：{approved_text_task.target}")
        print(f"  消息：{approved_text_task.message}")
        print(f"  可发送：{preview.get('sendable_count', 0)} 条")
        print(f"  模式：{'仅演练（不会发送）' if args.simulate else '真实发送'}")
        if not args.yes and not _confirm_send():
            print(json.dumps({"task_id": task_id, "status": "CANCELLED_BY_USER", "real_send": False}, ensure_ascii=False, indent=2))
            return
        assert_snapshot_unchanged(plan, snapshot)
        approval = approvals.issue(task_id, state.plan_hash or "")
        database.save_approval(approval)
        database.transition(task_id, "AUTHORIZED")
        result = _run_authorized(
            workspace=workspace,
            state_dir=state_dir,
            database=database,
            approvals=approvals,
            toolbox=toolbox,
            state=database.load_task(task_id),
            plan=plan,
            expected_snapshot=snapshot,
            approval_token=approval.token,
            simulate=args.simulate,
            approved_text_task=approved_text_task,
        )
        print(json.dumps({"task_id": task_id, **result}, ensure_ascii=False, indent=2))
        return

    if args.command == "preview":
        model = OpenAIResponsesClient(args.model) if args.live else RuleBasedMockClient()
        outcome = WeComAgent(model, toolbox).run(args.request)
        preview = outcome.get("latest_result")
        if not preview or preview.get("mode") != "preview":
            print(json.dumps(outcome, ensure_ascii=False, indent=2))
            raise SystemExit(2)
        plan = SendPlan.from_dict(preview["plan"]).resolve_paths(workspace)
        snapshot = capture_plan_snapshot(plan)
        task_id = f"task-{uuid.uuid4().hex[:10]}"
        state = TaskState(
            task_id=task_id,
            status="AWAITING_CONFIRMATION",
            plan=plan.as_dict(),
            plan_hash=plan.plan_hash(),
            preview=preview,
            snapshot=snapshot.as_dict(),
        )
        database.create_task(state)
        print(json.dumps({"task_id": task_id, **outcome, "snapshot_digest": snapshot.digest}, ensure_ascii=False, indent=2))
        return

    state = database.load_task(args.task_id)
    if args.command == "status":
        print(json.dumps({
            "task_id": state.task_id,
            "status": state.status,
            "plan_hash": state.plan_hash,
            "snapshot_digest": (state.snapshot or {}).get("digest"),
            "result": state.result,
            "audit_events": database.audit_events(state.task_id),
        }, ensure_ascii=False, indent=2))
        return
    plan = SendPlan.from_dict(state.plan or {}).resolve_paths(workspace)
    expected_snapshot = PlanSnapshot.from_dict(state.snapshot or {})
    if args.command == "approve":
        if state.status != "AWAITING_CONFIRMATION":
            raise SystemExit(f"Task is not awaiting confirmation: {state.status}")
        assert_snapshot_unchanged(plan, expected_snapshot)
        approval = approvals.issue(state.task_id, state.plan_hash or "")
        database.save_approval(approval)
        database.transition(state.task_id, "AUTHORIZED")
        print(json.dumps({
            "task_id": state.task_id,
            "status": "AUTHORIZED",
            "plan_hash": state.plan_hash,
            "expires_at": approval.expires_at,
            "approval_token": approval.token,
        }, ensure_ascii=False, indent=2))
        return

    if state.status != "AUTHORIZED":
        raise SystemExit(f"Task is not authorized: {state.status}")
    approved_text_task = None
    if args.command == "execute":
        policy = SupervisedExecutionPolicy(_allowed_targets(state_dir))
        approved_text_task = policy.validate(plan)
    result = _run_authorized(
        workspace=workspace,
        state_dir=state_dir,
        database=database,
        approvals=approvals,
        toolbox=toolbox,
        state=state,
        plan=plan,
        expected_snapshot=expected_snapshot,
        approval_token=args.approval_token,
        simulate=args.command == "simulate",
        approved_text_task=approved_text_task,
    )
    print(json.dumps({"task_id": state.task_id, **result}, ensure_ascii=False, indent=2))


def _run_authorized(
    *,
    workspace: Path,
    state_dir: Path,
    database: SQLiteStateStore,
    approvals: ApprovalManager,
    toolbox: WeComToolbox,
    state: TaskState,
    plan: SendPlan,
    expected_snapshot: PlanSnapshot,
    approval_token: str,
    simulate: bool,
    approved_text_task: ApprovedTextTask | None,
) -> dict[str, object]:
    approval = approvals.verify(state.task_id, state.plan_hash or "", approval_token)
    current_snapshot = assert_snapshot_unchanged(plan, expected_snapshot)
    idempotency_key = execution_idempotency_key(state.task_id, state.plan_hash or "", current_snapshot.digest)
    lock = ExecutionLock(state_dir / "execution.lock", state.task_id)
    with lock:
        mode = "simulation" if simulate else "supervised_real"
        execution_id = database.begin_execution_with_approval(approval, idempotency_key, mode)
        try:
            if simulate:
                result = toolbox.simulate_verified(plan)
            else:
                assert approved_text_task is not None
                result = RealWeComExecutionRunner(workspace).execute(approved_text_task)
            result["execution_id"] = execution_id
            result["idempotency_key"] = idempotency_key
            database.complete_execution(execution_id, "completed", result)
            database.transition(state.task_id, "COMPLETED", result=result)
        except Exception as exc:
            failure = {"error": str(exc), "execution_id": execution_id}
            database.complete_execution(execution_id, "failed", failure)
            database.transition(state.task_id, "FAILED", result=failure)
            raise
    return result


def _confirm_send() -> bool:
    try:
        answer = input("\n确认继续？[y/N] ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return False
    return answer in {"y", "yes"}


def _allowed_targets(state_dir: Path) -> tuple[str, ...]:
    targets = set(parse_allowed_targets(os.environ.get("WECOM_ALLOWED_TARGETS")))
    path = state_dir / "allowed-targets.txt"
    if path.is_file():
        targets.update(line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip())
    return tuple(sorted(targets))


def _approval_secret(state_dir: Path) -> str:
    configured = os.environ.get("WECOM_AGENT_APPROVAL_SECRET")
    if configured:
        return configured
    state_dir.mkdir(parents=True, exist_ok=True)
    path = state_dir / ".approval-secret"
    if path.exists():
        return path.read_text(encoding="utf-8").strip()
    secret = secrets.token_hex(32)
    path.write_text(secret, encoding="utf-8")
    try:
        path.chmod(0o600)
    except OSError:
        pass
    return secret


if __name__ == "__main__":
    main()
