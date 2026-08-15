from pathlib import Path
import json
import os
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from wecom_agent.execution_policy import (
    ExecutionPolicyDenied,
    SupervisedExecutionPolicy,
    parse_allowed_targets,
)
from wecom_agent.execution_runner import (
    FakeExecutionRunner,
    RealExecutionError,
    RealWeComExecutionRunner,
)
from wecom_agent.database import SQLiteStateStore
from wecom_agent.permissions import ApprovalManager
from wecom_agent.schemas import SendPlan
from wecom_agent.snapshots import capture_plan_snapshot
from wecom_agent.state_store import TaskState


HEADERS = [
    "渠道", "对象类型", "发送对象", "消息类型", "发送内容", "图片路径",
    "文件路径", "计划发送时间", "是否发送", "发送状态",
]


class SupervisedExecutionPolicyTest(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.workspace = Path(self.temporary.name)
        self.workbook = self.workspace / "smoke.xlsx"

    def tearDown(self):
        self.temporary.cleanup()

    def write_rows(self, *rows):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(HEADERS)
        for row in rows:
            sheet.append(row)
        workbook.save(self.workbook)

    def valid_row(self, **overrides):
        values = {
            "渠道": "企业微信",
            "对象类型": "个人",
            "发送对象": "授权测试对象",
            "消息类型": "文字",
            "发送内容": "Agent supervised send smoke test，请忽略。",
            "图片路径": "",
            "文件路径": "",
            "计划发送时间": "",
            "是否发送": "是",
            "发送状态": "",
        }
        values.update(overrides)
        return [values[header] for header in HEADERS]

    def plan(self, **overrides):
        value = {
            "workbook": str(self.workbook),
            "lesson_workbook": None,
            "target_workbook": None,
            "lesson": None,
            "selection": {"rows": [], "row_from": None, "row_to": None},
            "schedule_mode": "immediate",
            "run_at": None,
            "resend": False,
        }
        value.update(overrides)
        return SendPlan.from_dict(value).resolve_paths(self.workspace)

    def policy(self, *targets):
        return SupervisedExecutionPolicy(targets or ("授权测试对象",))

    def write_preview_sender_stub(self):
        sender = self.workspace / "scripts" / "send_from_excel_1v1_text.py"
        sender.parent.mkdir(exist_ok=True)
        sender.write_text(
            "import json\nprint(json.dumps({'sendable_rows': 1, 'skipped_rows': 0, "
            "'target_count': 1, 'batch_id': 'preview-test', 'run_dir': 'run-test'}))\n",
            encoding="utf-8",
        )

    def last_output_json(self, output):
        return json.loads(output[output.rfind("{\n"):])

    def test_allows_one_immediate_plain_text_row_for_exact_target(self):
        self.write_rows(self.valid_row())
        task = self.policy().validate(self.plan())
        self.assertEqual(task.row_number, 2)
        self.assertEqual(task.target, "授权测试对象")

    def test_requires_nonempty_exact_allowlist(self):
        self.write_rows(self.valid_row())
        with self.assertRaises(ExecutionPolicyDenied):
            SupervisedExecutionPolicy(()).validate(self.plan())
        with self.assertRaises(ExecutionPolicyDenied):
            self.policy("授权测试").validate(self.plan())

    def test_rejects_more_than_one_nonempty_selected_row(self):
        self.write_rows(self.valid_row(), self.valid_row())
        with self.assertRaises(ExecutionPolicyDenied):
            self.policy().validate(self.plan())

    def test_explicit_selection_cannot_hide_a_second_workbook_row(self):
        self.write_rows(self.valid_row(), self.valid_row(发送对象="另一测试对象"))
        with self.assertRaises(ExecutionPolicyDenied):
            self.policy().validate(self.plan(selection={"rows": [2], "row_from": None, "row_to": None}))

    def test_rejects_attachments(self):
        for field in ("图片路径", "文件路径"):
            with self.subTest(field=field):
                self.write_rows(self.valid_row(**{field: "unsafe.bin"}))
                with self.assertRaises(ExecutionPolicyDenied):
                    self.policy().validate(self.plan())

    def test_rejects_schedule_and_resend(self):
        self.write_rows(self.valid_row())
        with self.assertRaises(ExecutionPolicyDenied):
            self.policy().validate(self.plan(schedule_mode="scheduled", run_at="2030-01-01 09:00"))
        with self.assertRaises(ExecutionPolicyDenied):
            self.policy().validate(self.plan(resend=True))
        self.write_rows(self.valid_row(计划发送时间="2030-01-01 09:00"))
        with self.assertRaises(ExecutionPolicyDenied):
            self.policy().validate(self.plan())

    def test_rejects_group_non_text_disabled_and_prior_status(self):
        cases = (
            {"对象类型": "群聊"},
            {"消息类型": "图片"},
            {"是否发送": "否"},
            {"发送状态": "已发送"},
        )
        for overrides in cases:
            with self.subTest(overrides=overrides):
                self.write_rows(self.valid_row(**overrides))
                with self.assertRaises(ExecutionPolicyDenied):
                    self.policy().validate(self.plan())

    def test_fake_runner_never_claims_real_send(self):
        self.write_rows(self.valid_row())
        task = self.policy().validate(self.plan())
        result = FakeExecutionRunner().execute(task)
        self.assertFalse(result["real_send"])
        self.assertEqual(result["mode"], "supervised_fake")
        self.assertIn("not opened", result["message"])

    def test_real_runner_uses_exact_approved_row_and_requires_verified_success(self):
        self.write_rows(self.valid_row())
        task = self.policy().validate(self.plan())
        sender = self.workspace / "scripts" / "send_from_excel_1v1_text.py"
        sender.parent.mkdir()
        sender.write_text("# test sender\n", encoding="utf-8")
        calls = []

        def successful_runner(command, **kwargs):
            calls.append((command, kwargs))
            return subprocess.CompletedProcess(
                command, 0, stdout=json.dumps({
                    "success": 1, "late_success": 0, "failed": 0,
                    "batch_id": "batch-test", "run_dir": "run-test",
                }), stderr="",
            )

        result = RealWeComExecutionRunner(self.workspace, successful_runner).execute(task)
        command = calls[0][0]
        self.assertEqual(result["mode"], "supervised_real")
        self.assertTrue(result["real_send"])
        self.assertEqual(command[command.index("--row") + 1], "2")
        self.assertIn("--execute", command)
        self.assertIn("--yes", command)
        self.assertIn("--save-each-row", command)
        self.assertIn("--stop-on-error", command)

        def unverified_runner(command, **kwargs):
            return subprocess.CompletedProcess(
                command, 0, stdout=json.dumps({"success": 0, "late_success": 0, "failed": 1}), stderr="",
            )

        with self.assertRaises(RealExecutionError):
            RealWeComExecutionRunner(self.workspace, unverified_runner).execute(task)

    def test_allowlist_parser_trims_and_drops_empty_values(self):
        self.assertEqual(parse_allowed_targets(" 测试甲,测试乙, ,"), ("测试甲", "测试乙"))

    def test_cli_policy_denial_does_not_consume_approval_then_real_execute_succeeds(self):
        self.write_rows(self.valid_row())
        plan = self.plan()
        snapshot = capture_plan_snapshot(plan)
        task = TaskState(
            task_id="task-supervised",
            status="AUTHORIZED",
            plan=plan.as_dict(),
            plan_hash=plan.plan_hash(),
            preview={"mode": "preview"},
            snapshot=snapshot.as_dict(),
        )
        state_dir = self.workspace / ".agent-state"
        state_dir.mkdir()
        secret = "integration-test-secret"
        (state_dir / ".approval-secret").write_text(secret, encoding="utf-8")
        database = SQLiteStateStore(state_dir / "agent.db")
        database.create_task(task)
        approval = ApprovalManager(secret).issue(task.task_id, task.plan_hash or "")
        database.save_approval(approval)

        sender = self.workspace / "scripts" / "send_from_excel_1v1_text.py"
        sender.parent.mkdir()
        sender.write_text(
            "import json\nprint(json.dumps({'success': 1, 'late_success': 0, 'failed': 0, 'batch_id': 'test'}))\n",
            encoding="utf-8",
        )

        command = [
            sys.executable, "-m", "wecom_agent.cli", "--workspace", str(self.workspace),
            "execute", task.task_id, "--approval-token", approval.token,
        ]
        environment = os.environ.copy()
        environment["PYTHONPATH"] = str(ROOT / "src")
        environment["WECOM_ALLOWED_TARGETS"] = "wrong-target"
        denied = subprocess.run(command, cwd=ROOT, env=environment, text=True, capture_output=True, check=False)
        self.assertNotEqual(denied.returncode, 0)
        self.assertEqual(database.load_task(task.task_id).status, "AUTHORIZED")

        environment["WECOM_ALLOWED_TARGETS"] = "授权测试对象"
        completed = subprocess.run(command, cwd=ROOT, env=environment, text=True, capture_output=True, check=False)
        self.assertEqual(completed.returncode, 0, completed.stderr)
        result = json.loads(completed.stdout)
        self.assertTrue(result["real_send"])
        self.assertEqual(result["mode"], "supervised_real")
        self.assertEqual(database.load_task(task.task_id).status, "COMPLETED")

    def test_one_click_send_simulates_without_gui_and_hides_approval_token(self):
        self.write_rows(self.valid_row())
        self.write_preview_sender_stub()
        command = [
            sys.executable, "-m", "wecom_agent.cli", "--workspace", str(self.workspace),
            "send", str(self.workbook), "--simulate", "--yes",
        ]
        environment = os.environ.copy()
        environment["PYTHONPATH"] = str(ROOT / "src")
        environment["WECOM_ALLOWED_TARGETS"] = "授权测试对象"
        completed = subprocess.run(command, cwd=ROOT, env=environment, text=True, capture_output=True, check=False)
        self.assertEqual(completed.returncode, 0, completed.stderr)
        payload = self.last_output_json(completed.stdout)
        self.assertEqual(payload["mode"], "simulation")
        self.assertFalse(payload.get("real_send", False))
        self.assertNotIn("approval_token", completed.stdout)
        database = SQLiteStateStore(self.workspace / ".agent-state" / "agent.db")
        self.assertEqual(database.load_task(payload["task_id"]).status, "COMPLETED")

    def test_allow_command_persists_local_target_for_one_click_send(self):
        self.write_rows(self.valid_row())
        self.write_preview_sender_stub()
        base = [sys.executable, "-m", "wecom_agent.cli", "--workspace", str(self.workspace)]
        environment = os.environ.copy()
        environment["PYTHONPATH"] = str(ROOT / "src")
        environment.pop("WECOM_ALLOWED_TARGETS", None)
        configured = subprocess.run(
            [*base, "allow", "授权测试对象"], cwd=ROOT, env=environment,
            text=True, capture_output=True, check=False,
        )
        self.assertEqual(configured.returncode, 0, configured.stderr)
        completed = subprocess.run(
            [*base, "send", str(self.workbook), "--simulate", "--yes"], cwd=ROOT, env=environment,
            text=True, capture_output=True, check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        payload = self.last_output_json(completed.stdout)
        self.assertEqual(payload["mode"], "simulation")

    def test_one_click_send_defaults_to_cancel_when_not_confirmed(self):
        self.write_rows(self.valid_row())
        self.write_preview_sender_stub()
        command = [
            sys.executable, "-m", "wecom_agent.cli", "--workspace", str(self.workspace),
            "send", str(self.workbook),
        ]
        environment = os.environ.copy()
        environment["PYTHONPATH"] = str(ROOT / "src")
        environment["WECOM_ALLOWED_TARGETS"] = "授权测试对象"
        completed = subprocess.run(
            command, cwd=ROOT, env=environment, input="n\n", text=True, capture_output=True, check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        payload = self.last_output_json(completed.stdout)
        self.assertEqual(payload["status"], "CANCELLED_BY_USER")
        self.assertFalse(payload["real_send"])


if __name__ == "__main__":
    unittest.main()
