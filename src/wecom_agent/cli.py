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
from .execution_policy import SupervisedExecutionPolicy, parse_allowed_targets
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
    approval = approvals.verify(state.task_id, state.plan_hash or "", args.approval_token)
    current_snapshot = assert_snapshot_unchanged(plan, expected_snapshot)
    approved_text_task = None
    if args.command == "execute":
        policy = SupervisedExecutionPolicy(parse_allowed_targets(os.environ.get("WECOM_ALLOWED_TARGETS")))
        approved_text_task = policy.validate(plan)
    idempotency_key = execution_idempotency_key(state.task_id, state.plan_hash or "", current_snapshot.digest)
    lock = ExecutionLock(state_dir / "execution.lock", state.task_id)
    with lock:
        mode = "simulation" if args.command == "simulate" else "supervised_real"
        execution_id = database.begin_execution_with_approval(approval, idempotency_key, mode)
        try:
            if args.command == "simulate":
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
    print(json.dumps({"task_id": state.task_id, **result}, ensure_ascii=False, indent=2))


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
