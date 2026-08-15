from datetime import datetime, timedelta, timezone
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from wecom_agent.agent import AgentLoopError, WeComAgent
from wecom_agent.database import DuplicateExecution, SQLiteStateStore
from wecom_agent.execution_lock import ExecutionLock, ExecutionLocked
from wecom_agent.model_client import ModelTurn, RuleBasedMockClient, ToolCall
from wecom_agent.permissions import ApprovalManager, PermissionDenied
from wecom_agent.schemas import PlanValidationError, SendPlan
from wecom_agent.snapshots import (
    SnapshotMismatch,
    assert_snapshot_unchanged,
    capture_plan_snapshot,
    delivery_key,
    execution_idempotency_key,
)
from wecom_agent.state_store import TaskState
from wecom_agent.tools import WeComToolbox


def direct_plan(**overrides):
    value = {
        "workbook": "examples/example_tasks.xlsx",
        "lesson_workbook": None,
        "target_workbook": None,
        "lesson": None,
        "selection": {"rows": [], "row_from": None, "row_to": None},
        "schedule_mode": "immediate",
        "run_at": None,
        "resend": False,
    }
    value.update(overrides)
    return value


def stored_task(task_id: str = "task-1") -> TaskState:
    plan = SendPlan.from_dict(direct_plan()).resolve_paths(ROOT)
    snapshot = capture_plan_snapshot(plan)
    return TaskState(
        task_id=task_id,
        status="AWAITING_CONFIRMATION",
        plan=plan.as_dict(),
        plan_hash=plan.plan_hash(),
        preview={"mode": "preview"},
        snapshot=snapshot.as_dict(),
    )


class AgentMvpTest(unittest.TestCase):
    def test_plan_hash_is_stable(self):
        first = SendPlan.from_dict(direct_plan())
        second = SendPlan.from_dict(direct_plan())
        self.assertEqual(first.plan_hash(), second.plan_hash())

    def test_plan_rejects_unsafe_path(self):
        plan = SendPlan.from_dict(direct_plan(workbook="../private.xlsx"))
        with self.assertRaises(PlanValidationError):
            plan.resolve_paths(ROOT)

    def test_plan_rejects_conflicting_row_selection(self):
        with self.assertRaises(PlanValidationError):
            SendPlan.from_dict(direct_plan(selection={"rows": [2], "row_from": 2, "row_to": 3}))

    def test_approval_is_bound_to_task_plan_and_expiry(self):
        manager = ApprovalManager("test-secret")
        now = datetime(2030, 1, 1, tzinfo=timezone.utc)
        plan = SendPlan.from_dict(direct_plan())
        approval = manager.issue("task-1", plan.plan_hash(), now=now, ttl_seconds=60)
        manager.verify("task-1", plan.plan_hash(), approval.token, now=now + timedelta(seconds=30))
        with self.assertRaises(PermissionDenied):
            manager.verify("task-2", plan.plan_hash(), approval.token, now=now)
        with self.assertRaises(PermissionDenied):
            manager.verify("task-1", plan.plan_hash(), approval.token, now=now + timedelta(seconds=60))

    def test_state_machine_rejects_execute_before_approval(self):
        state = TaskState("task-1")
        with self.assertRaises(ValueError):
            state.transition("EXECUTING")

    def test_mock_agent_runs_preview_tool(self):
        toolbox = WeComToolbox(ROOT)
        outcome = WeComAgent(RuleBasedMockClient(), toolbox).run("预览示例任务")
        self.assertEqual(outcome["latest_result"]["mode"], "preview")
        self.assertIn("预览完成", outcome["message"])

    def test_simulation_is_internal_and_never_runs_subprocess(self):
        calls = []

        def runner(*args, **kwargs):
            calls.append(args)
            return subprocess.CompletedProcess(args[0], 0, "{}", "")

        toolbox = WeComToolbox(ROOT, command_runner=runner)
        plan = SendPlan.from_dict(direct_plan()).resolve_paths(ROOT)
        result = toolbox.simulate_verified(plan)
        self.assertEqual(result["mode"], "simulation")
        self.assertEqual(calls, [])
        self.assertEqual([item.name for item in toolbox.definitions], ["preview_wecom_tasks"])

    def test_unknown_tool_is_rejected(self):
        toolbox = WeComToolbox(ROOT)
        with self.assertRaises(Exception):
            toolbox.dispatch("shell", {"command": "echo unsafe"})

    def test_agent_stops_repeated_tool_calls(self):
        class LoopingModel:
            def start(self, user_input, tools):
                return ModelTurn(tool_calls=(ToolCall("loop-0", "preview_wecom_tasks", direct_plan()),))

            def continue_with(self, outputs, tools):
                return ModelTurn(tool_calls=(ToolCall("loop-next", "preview_wecom_tasks", direct_plan()),))

        with self.assertRaises(AgentLoopError):
            WeComAgent(LoopingModel(), WeComToolbox(ROOT), max_tool_rounds=1).run("loop")


class ProductionSafetyTest(unittest.TestCase):
    def test_sqlite_approval_can_only_be_consumed_once(self):
        with tempfile.TemporaryDirectory() as temporary:
            database = SQLiteStateStore(Path(temporary) / "agent.db")
            state = stored_task()
            database.create_task(state)
            manager = ApprovalManager("test-secret")
            approval = manager.issue(state.task_id, state.plan_hash or "", ttl_seconds=60)
            database.save_approval(approval)
            verified = manager.verify(state.task_id, state.plan_hash or "", approval.token)
            database.consume_approval(verified)
            with self.assertRaises(PermissionDenied):
                database.consume_approval(verified)

    def test_sqlite_records_audit_transitions(self):
        with tempfile.TemporaryDirectory() as temporary:
            database = SQLiteStateStore(Path(temporary) / "agent.db")
            state = stored_task()
            database.create_task(state)
            database.transition(state.task_id, "AUTHORIZED")
            event_types = [event["type"] for event in database.audit_events(state.task_id)]
            self.assertEqual(event_types, ["task_created", "status_changed"])

    def test_begin_execution_atomically_consumes_approval_and_changes_state(self):
        with tempfile.TemporaryDirectory() as temporary:
            database = SQLiteStateStore(Path(temporary) / "agent.db")
            state = stored_task()
            database.create_task(state)
            database.transition(state.task_id, "AUTHORIZED")
            manager = ApprovalManager("test-secret")
            approval = manager.issue(state.task_id, state.plan_hash or "")
            database.save_approval(approval)
            verified = manager.verify(state.task_id, state.plan_hash or "", approval.token)
            key = execution_idempotency_key(state.task_id, state.plan_hash or "", state.snapshot["digest"])
            execution_id = database.begin_execution_with_approval(verified, key, "simulation")
            self.assertTrue(execution_id)
            self.assertEqual(database.load_task(state.task_id).status, "SIMULATING")
            with self.assertRaises(PermissionDenied):
                database.begin_execution_with_approval(verified, key, "simulation")

    def test_database_rejects_expired_approval_even_without_manager_check(self):
        with tempfile.TemporaryDirectory() as temporary:
            database = SQLiteStateStore(Path(temporary) / "agent.db")
            state = stored_task()
            database.create_task(state)
            database.transition(state.task_id, "AUTHORIZED")
            manager = ApprovalManager("test-secret")
            issued = datetime(2030, 1, 1, tzinfo=timezone.utc)
            approval = manager.issue(state.task_id, state.plan_hash or "", now=issued, ttl_seconds=5)
            database.save_approval(approval)
            key = execution_idempotency_key(state.task_id, state.plan_hash or "", state.snapshot["digest"])
            with self.assertRaises(PermissionDenied):
                database.begin_execution_with_approval(
                    approval,
                    key,
                    "simulation",
                    now=issued + timedelta(seconds=5),
                )

    def test_execution_idempotency_rejects_duplicate_start(self):
        with tempfile.TemporaryDirectory() as temporary:
            database = SQLiteStateStore(Path(temporary) / "agent.db")
            state = stored_task()
            database.create_task(state)
            key = execution_idempotency_key(state.task_id, state.plan_hash or "", state.snapshot["digest"])
            database.start_execution(state.task_id, key, "simulation")
            with self.assertRaises(DuplicateExecution):
                database.start_execution(state.task_id, key, "simulation")

    def test_delivery_key_is_stable_and_content_sensitive(self):
        first = delivery_key("plan", "snapshot", 2, "示例用户甲", "message-a")
        second = delivery_key("plan", "snapshot", 2, "示例用户甲", "message-a")
        changed = delivery_key("plan", "snapshot", 2, "示例用户甲", "message-b")
        self.assertEqual(first, second)
        self.assertNotEqual(first, changed)

    def test_workbook_change_invalidates_snapshot(self):
        with tempfile.TemporaryDirectory() as temporary:
            workspace = Path(temporary)
            examples = workspace / "examples"
            examples.mkdir()
            workbook_path = examples / "example_tasks.xlsx"
            shutil.copy2(ROOT / "examples" / "example_tasks.xlsx", workbook_path)
            plan = SendPlan.from_dict(direct_plan()).resolve_paths(workspace)
            snapshot = capture_plan_snapshot(plan)
            with workbook_path.open("ab") as handle:
                handle.write(b"changed-after-preview")
            with self.assertRaises(SnapshotMismatch):
                assert_snapshot_unchanged(plan, snapshot)

    def test_execution_lock_blocks_second_owner_and_preserves_owner(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "execution.lock"
            first = ExecutionLock(path, "task-1")
            second = ExecutionLock(path, "task-2")
            first.acquire()
            try:
                with self.assertRaises(ExecutionLocked):
                    second.acquire()
                self.assertTrue(path.exists())
            finally:
                first.release()
            self.assertFalse(path.exists())


if __name__ == "__main__":
    unittest.main()
