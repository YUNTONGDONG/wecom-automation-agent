---
name: wecom-message-automation
description: Execute safety-checked WeCom message delivery tasks from local Excel workbooks, including individual or group targets, scheduled sending, text and attachments, OCR-based target and input verification, post-send validation, Excel write-back, retry control, and evidence logging. Use when Codex must preview, validate, schedule, send, resume, or audit structured Enterprise WeChat/WeCom messaging tasks through the desktop client.
---

# WeCom Message Automation

Use `scripts/send_from_excel_1v1_text.py` as the deterministic execution engine. Treat desktop messaging as a high-risk external action: preview first, preserve target checks, and require explicit authorization before sending.

## Workflow

1. Locate the source workbook and resolve all paths relative to `--folder`.
2. Run preview mode and inspect every eligible row, target, message type, attachment path, schedule, and skip reason.
3. Report the intended target count and any invalid or skipped rows. Do not infer missing recipients or repair ambiguous names automatically.
4. Obtain explicit user confirmation for the reviewed plan before executing. A request that already clearly says to send the specified workbook/rows counts as authorization; otherwise stop after preview.
5. Execute one batch process with `--execute --yes`. Reuse `--run-dir` and `--batch-id` if preview and execution belong to the same user request.
6. Inspect `result.json`, `evidence_manifest.json`, `run_log.txt`, and Excel write-back fields. Report success, late sends, failures, skipped rows, and rows requiring manual verification separately.

## Commands

Preview a prepared send workbook:

```bash
python scripts/send_from_excel_1v1_text.py \
  --folder /absolute/project/folder \
  --workbook /absolute/path/tasks.xlsx \
  --json
```

Execute the confirmed plan:

```bash
python scripts/send_from_excel_1v1_text.py \
  --folder /absolute/project/folder \
  --workbook /absolute/path/tasks.xlsx \
  --execute --yes --json
```

Generate and send a lesson/reminder workbook in one process:

```bash
python scripts/send_from_excel_1v1_text.py \
  --folder /absolute/project/folder \
  --lesson-workbook /absolute/path/content.xlsx \
  --target-workbook /absolute/path/targets.xlsx \
  --lesson 3 \
  --respect-send-time \
  --execute --yes --json
```

Use `--row N` repeatedly for selected rows, or `--row-from N --row-to M` for a range. Add `--resend` only when the user explicitly requests duplicate delivery of rows already marked sent. Preserve Excel order unless the user accepts `--group-targets`.

## Input contract

Recognize these field groups:

- Routing: `渠道`, `对象类型`, `发送对象` and supported contact/group aliases.
- Target verification: `目标别名`, `核对关键词`, `会话关键词`, `英文名`, `别名`.
- Content: `发送内容` and supported text aliases; `图片路径`; `文件路径`; `消息类型`.
- Control: `计划发送时间`, `是否发送`, `发送状态`, `错误原因`, `发送时间`.

Keep one semantic message per row. When one row contains text plus attachments, send them together in one action. Split content into multiple rows only when the user intends separate message bubbles.

Reject or skip rows with an unsupported channel, missing target, missing required attachment, a false send flag, or an existing sent status unless `--resend` is authorized. Never silently substitute a similar target.

## Safety invariants

- Require both `--execute` and `--yes` for a real send.
- Confirm the foreground application is WeCom before every click, paste, or shortcut.
- Verify the search result and opened conversation against the exact target or an explicit alias. A longer prefix match is not an exact match.
- Verify the clipboard and visible input content before clicking Send. Clear unrelated or residual text and stop that row when verification fails.
- Never bypass target, search, input, or post-send checks merely for speed. Use skip or experimental fast flags only for a user-approved, supervised diagnostic run.
- For scheduled tasks, use `--respect-send-time`; preparation may start early, but the send action must not occur before the planned time.
- Stop or quarantine the affected row when WeCom shows a security check, environment restriction, failure marker, or unresolved upload state.
- Do not run two sender processes concurrently against the same desktop client.

## Validation and evidence

Accept success only when a new outgoing message is visible and the newest message matches the current row. Do not use older bubbles, warnings, or timestamps as evidence. Treat unresolved link parsing or ambiguous UI state as `发送待核对`, not success.

Keep one run directory under `wechat-work-message-validation/logs/` for the whole request. Expect:

- `run_log.txt` for concise milestones and row outcomes.
- `result.json` for structured execution results.
- `evidence_manifest.json` for real-send evidence references.
- Row-level screenshots and OCR artifacts when evidence capture is enabled.
- Excel write-back fields for checks, timestamps, batch ID, status, errors, and duration.

Do not commit logs, screenshots, generated send workbooks, real contact workbooks, or clipboard/message contents to source control.

## Platform notes

Support macOS and Windows desktop automation through the implementation adapters. On macOS, grant Accessibility and Screen Recording permissions to the terminal/Codex host. On Windows, keep the bundled PowerShell OCR helper available. Perform a preview and a supervised 1-3 row smoke test after any WeCom UI update, display-scaling change, or OCR/runtime change.
