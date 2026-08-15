import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "send_from_excel_1v1_text.py"
SPEC = importlib.util.spec_from_file_location("sender", SCRIPT)
sender = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = sender
SPEC.loader.exec_module(sender)


def row(
    message_type: str,
    message: str = "",
    images: list[str] | None = None,
    docs: list[str] | None = None,
) -> sender.SendRow:
    images = images or []
    docs = docs or []
    parts = sender.parse_parts(message_type, message, images, docs)
    return sender.SendRow(
        row=2,
        channel="企业微信",
        object_type="个人",
        target="示例用户甲",
        target_aliases=["Demo User A"],
        scheduled_at="",
        message=message,
        message_type=message_type or "自动",
        image_paths=[Path(path) for path in images],
        document_paths=[Path(path) for path in docs],
        parts=parts,
    )


def route_for(item: sender.SendRow) -> tuple[list[str], list[str]]:
    gui = sender.WeComGui.__new__(sender.WeComGui)
    gui.capture_evidence = False
    gui.last_rect = object()
    gui.between_rows = 0
    calls: list[str] = []

    def send_text(_rect, _item):
        calls.append("send_text")

    def send_text_with_files(_rect, _item, paths):
        calls.append(f"send_text_with_files:{len(paths)}")

    def send_files(_rect, _item, paths):
        calls.append(f"send_files:{len(paths)}")

    gui.activate = lambda: object()
    gui.send_text = send_text
    gui.send_text_with_files = send_text_with_files
    gui.send_files = send_files
    gui.wait_for_send_settle = lambda _item, _rect, _before: calls.append("settle")
    _seconds, steps = sender.WeComGui.send_message(gui, item)
    return calls, steps


class MessageMatrixTest(unittest.TestCase):
    def test_expected_steps_for_single_content(self):
        self.assertEqual(sender.expected_steps(row("文字", "hello")), ["文字"])
        self.assertEqual(sender.expected_steps(row("图片", images=["a.png"])), ["图片(1)"])
        self.assertEqual(sender.expected_steps(row("文档", docs=["a.docx"])), ["文档(1)"])

    def test_expected_steps_for_combined_content(self):
        self.assertEqual(sender.expected_steps(row("文字、图片", "hello", images=["a.png"])), ["文字+附件"])
        self.assertEqual(sender.expected_steps(row("文字、文档", "hello", docs=["a.docx"])), ["文字+附件"])
        self.assertEqual(
            sender.expected_steps(row("文字、图片、文档", "hello", images=["a.png"], docs=["a.docx"])),
            ["文字+附件"],
        )

    def test_expected_steps_for_non_text_mixed_attachments(self):
        self.assertEqual(
            sender.expected_steps(row("图片、文档", images=["a.png"], docs=["a.docx"])),
            ["图片(1)", "文档(1)"],
        )

    def test_send_route_for_text_with_attachment_is_one_combined_send(self):
        calls, steps = route_for(row("文字、图片", "hello", images=["a.png"]))
        self.assertEqual(calls, ["send_text_with_files:1", "settle"])
        self.assertEqual(steps, ["文字+附件"])

    def test_send_route_for_separate_attachment_only_remains_attachment_send(self):
        calls, steps = route_for(row("图片", images=["a.png"]))
        self.assertEqual(calls, ["send_files:1", "settle"])
        self.assertEqual(steps, ["图片(1)"])

    def test_fast_dispatch_send_does_not_wait_for_settle(self):
        item = row("文字", "hello")
        gui = sender.WeComGui.__new__(sender.WeComGui)
        gui.capture_evidence = False
        gui.last_rect = object()
        gui.between_rows = 0
        calls: list[str] = []
        gui.activate = lambda: object()
        gui.send_text = lambda _rect, _item: calls.append("send_text")
        gui.send_text_with_files = lambda _rect, _item, _paths: calls.append("send_text_with_files")
        gui.send_files = lambda _rect, _item, _paths: calls.append("send_files")
        gui.wait_for_send_settle = lambda _item, _rect, _before: calls.append("settle")
        _seconds, steps = sender.WeComGui.send_message(gui, item, wait_settle=False)
        self.assertEqual(calls, ["send_text"])
        self.assertEqual(steps, ["文字"])

    def test_parse_parts_infers_type_when_message_type_is_blank(self):
        item = row("", "hello", images=["a.png"], docs=["a.docx"])
        self.assertEqual(item.parts, {"text": True, "image": True, "file": True})
        self.assertEqual(sender.expected_steps(item), ["文字+附件"])

    def test_input_text_capture_uses_window_relative_crop(self):
        from PIL import Image, ImageDraw

        with tempfile.TemporaryDirectory() as tmp:
            gui = sender.WeComGui.__new__(sender.WeComGui)
            gui.run_dir = Path(tmp)
            gui.input_panel_bbox = lambda _rect: (rect.left + 425, rect.bottom - 160, rect.right - 24, rect.bottom - 16)
            rect = sender.WinRect()
            rect.left = 704
            rect.top = 147
            rect.right = 1790
            rect.bottom = 797
            item = row("文字", "hello")
            window_bbox = (rect.left, rect.top, rect.right, rect.bottom)

            def fake_capture_bbox(bbox, path):
                path.parent.mkdir(parents=True, exist_ok=True)
                if bbox == window_bbox:
                    width = (rect.right - rect.left) * 2
                    height = (rect.bottom - rect.top) * 2
                    image = Image.new("RGB", (width, height), (20, 40, 180))
                    input_top = int((rect.bottom - rect.top - 165) * 2)
                    ImageDraw.Draw(image).rectangle((0, input_top, width, height), fill=(20, 180, 60))
                    image.save(path)
                else:
                    Image.new("RGB", (120, 80), (220, 20, 20)).save(path)
                return path

            gui.capture_bbox = fake_capture_bbox
            output = sender.WeComGui.capture_input_text(gui, item, rect)
            cropped = Image.open(output).convert("RGB")
            pixels = list(cropped.get_flattened_data())
            greenish = sum(1 for red, green, blue in pixels if green > red + 80 and green > blue + 20)
            self.assertGreater(greenish, len(pixels) * 0.8)

    def test_detect_input_panel_ignores_group_sidebar(self):
        from PIL import Image, ImageDraw

        image = Image.new("RGB", (1086, 650), (245, 247, 250))
        draw = ImageDraw.Draw(image)
        draw.rectangle((427, 492, 907, 634), fill=(255, 255, 255))
        draw.rectangle((926, 0, 1085, 650), fill=(255, 255, 255))
        bbox = sender.detect_input_panel_bbox_from_image(image)
        self.assertIsNotNone(bbox)
        self.assertEqual(bbox[0], 427)
        self.assertEqual(bbox[2], 907)

    def test_target_ocr_variants_for_zhe_character(self):
        self.assertIn("测试吉", sender.target_ocr_variants("测试喆"))
        self.assertEqual(sender.target_ocr_variants("甲乙"), ["甲乙"])

    def test_target_ocr_variants_for_shen_preserve_short_name_safety(self):
        candidates = sender.target_ocr_candidates("合成珅珅", [], include_variants=True)
        self.assertIn("合成呻呻", candidates)
        self.assertIn(sender.ocr_contains_any_target(candidates, "合 成 呻 呻 合 成 砷 砷"), {"合成呻呻", "合成砷砷"})

        short_candidates = sender.target_ocr_candidates("合成珅", [], include_variants=True)
        self.assertEqual(sender.ocr_contains_any_target(short_candidates, "合 成 呻 呻"), "")

    def test_settle_accepts_matching_text_when_bubble_position_is_reused(self):
        from PIL import Image, ImageDraw

        message = "课程链接：https://example.invalid/course/lesson-03"
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            gui = sender.WeComGui.__new__(sender.WeComGui)
            gui.capture_evidence = True
            gui.run_dir = tmp_path
            gui.send_settle_timeout = 0.05
            gui.send_settle_interval = 0.01
            item = row("文字", message)
            item.send_action_time = "2026-06-02 17:25:23"

            rect = sender.WinRect()
            rect.left = 0
            rect.top = 0
            rect.right = 900
            rect.bottom = 600

            def save_window(path: Path) -> Path:
                image = Image.new("RGB", (900, 600), (245, 247, 250))
                draw = ImageDraw.Draw(image)
                draw.rectangle((305, 90, 850, 430), fill=(200, 232, 255))
                image.save(path)
                return path

            before = save_window(tmp_path / "before.png")

            def capture_window(_item, _rect, stage):
                path = tmp_path / f"{stage}.png"
                save_window(path)
                if stage == "after_send":
                    _item.after_screenshot = str(path)
                return path

            def blank_evidence(stage: str) -> Path:
                path = tmp_path / f"{stage}.png"
                Image.new("RGB", (260, 120), (255, 255, 255)).save(path)
                return path

            gui.capture_window = capture_window
            gui.capture_send_status = lambda _item, _rect: blank_evidence("send_status")
            gui.capture_input_text = lambda _item, _rect: blank_evidence("input_text")
            gui.ocr_image = lambda path: message if "latest_message" in Path(path).name else ""

            sender.WeComGui.wait_for_send_settle(gui, item, rect, before)
            self.assertEqual(item.step_events[-1]["status"], "已通过")

    def test_input_message_match_accepts_visible_middle_of_long_draft(self):
        message = (
            "示例自动化课程提醒，请核对发送对象和消息内容。\n"
            "测试专题第一课：可靠消息投递。\n"
            "执行前要完成三个检查：\n"
            "1.核对发送对象\n"
            "2.核对消息内容\n"
            "3.核对发送结果\n"
            "课程链接：https://example.invalid/course/lesson-03?scene=test&source=demo"
        )
        visible_ocr = "1 · 核 对 发 送 对 象 2 · 核 对 消 息 内 容 3 · 核 对 发 送 结 果"
        self.assertTrue(sender.ocr_matches_input_message(message, visible_ocr))

    def test_input_message_match_rejects_unrelated_draft(self):
        message = "示例自动化课程提醒，请核对发送对象和消息内容。\n课程链接：https://example.invalid/course/lesson-03"
        self.assertFalse(sender.ocr_matches_input_message(message, "当前设备环境异常，需要扫码验证"))

    def test_default_text_paste_methods_are_keyboard_retries_only(self):
        gui = sender.WeComGui.__new__(sender.WeComGui)
        gui.paste_method_order = sender.DEFAULT_PASTE_METHOD_ORDER
        names = [name for name, _method in sender.WeComGui.paste_methods(gui, 100, 100)]
        self.assertEqual(names, ["Ctrl+V", "Ctrl+V", "Ctrl+V"])

    def test_text_paste_methods_do_not_auto_append_unsafe_fallbacks(self):
        gui = sender.WeComGui.__new__(sender.WeComGui)
        gui.paste_method_order = "ctrl-v,wm-point"
        names = [name for name, _method in sender.WeComGui.paste_methods(gui, 100, 100)]
        self.assertEqual(names, ["Ctrl+V", "WM_PASTE输入框"])

    def test_wecom_unavailable_contact_error_marker(self):
        self.assertEqual(sender.send_error_marker("示例用户甲已离开当前企业，无法发送消息"), "无法发送")


if __name__ == "__main__":
    unittest.main()
