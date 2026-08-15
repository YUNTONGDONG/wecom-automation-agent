# WeCom Automation Agent

A stateful desktop automation agent for reliable WeCom messaging from structured Excel tasks. It combines target verification, scheduled execution, message composition, post-action validation, Excel write-back, and evidence-backed audit logs.

> This project controls a real messaging client. Start with preview mode and a supervised 1-3 row test. Never use production contact data in a public repository.

## Why this is more than a send script

The execution pipeline treats every GUI action as untrusted until it is verified:

```text
Excel task
  -> eligibility and path validation
  -> exact target search and OCR confirmation
  -> clipboard and input-box verification
  -> scheduled or immediate send
  -> newest-message validation
  -> Excel write-back + JSON result + evidence manifest
```

The implementation supports text, images, documents, combined messages, individual contacts, group chats, row selection, scheduled delivery, retry control, resuming, and macOS/Windows OCR adapters.

## Repository layout

```text
wecom-fixed-message-draft/
├── SKILL.md                         # Agent-facing trigger and operating procedure
├── README.md                        # Human-facing setup and usage
├── requirements.txt
├── .gitignore
├── agents/openai.yaml              # Codex UI metadata
├── docs/architecture.md
├── examples/example_tasks.xlsx     # Synthetic, non-production input
├── scripts/
│   ├── send_from_excel_1v1_text.py # Execution engine
│   ├── excel_utils.py
│   ├── log_utils.py
│   ├── macos_vision_ocr.py
│   └── ocr_windows_image.ps1
└── tests/
```

The large execution engine is intentionally retained for this first packaged release to avoid changing verified GUI behavior. Pure parsing and validation helpers can be extracted incrementally behind tests.

## Requirements

- Python 3.10+
- WeCom/Enterprise WeChat desktop client, logged in and visible
- Microsoft Excel-compatible `.xlsx` task workbook
- Accessibility and Screen Recording permission on macOS
- A Windows desktop session with PowerShell OCR support on Windows

Create a virtual environment and install dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

On Windows activate with `.venv\Scripts\activate`.

## Quick start

Preview the synthetic example without sending anything:

```bash
python scripts/send_from_excel_1v1_text.py \
  --folder . \
  --workbook examples/example_tasks.xlsx \
  --json
```

For a real workbook, review preview output first. Then run the confirmed batch:

```bash
python scripts/send_from_excel_1v1_text.py \
  --folder /absolute/project/folder \
  --workbook /absolute/path/tasks.xlsx \
  --execute --yes --json
```

Use selected rows:

```bash
python scripts/send_from_excel_1v1_text.py \
  --folder . --workbook /absolute/path/tasks.xlsx \
  --row 2 --row 5 --execute --yes --json
```

Respect planned send times:

```bash
python scripts/send_from_excel_1v1_text.py \
  --folder . --workbook /absolute/path/tasks.xlsx \
  --respect-send-time --execute --yes --json
```

Run `python scripts/send_from_excel_1v1_text.py --help` for advanced execution, evidence, row-range, and diagnostic flags.

## Workbook schema

The simplest task sheet uses these columns:

| Column | Purpose |
|---|---|
| `渠道` | `企业微信`, `企微`, `WXWork`, or `WeCom` |
| `对象类型` | Individual contact or group chat |
| `发送对象` | Exact visible contact/group name |
| `目标别名` | Optional explicit OCR-safe aliases |
| `消息类型` | Text, image, document, or a combination |
| `发送内容` | Text body |
| `图片路径` | Image path(s), when required |
| `文件路径` | Document path(s), when required |
| `计划发送时间` | Optional date/time |
| `是否发送` | Boolean-like execution flag |
| `发送状态` | Written back by the program |

Several Chinese aliases are accepted for compatibility. Keep one intended message bubble per row. A row containing text and attachments is sent as one combined message; use separate rows for separate bubbles.

## Safety model

- Real delivery requires the two-part gate `--execute --yes`.
- Exact target/alias matching blocks prefix collisions such as a short name matching a longer one.
- Search, conversation title, clipboard, and input content are checked before sending.
- The newest outgoing bubble is checked after sending; old messages do not count as proof.
- Security prompts, red failure markers, unresolved uploads, and ambiguous outcomes are not written back as successful.
- Existing successful rows are skipped unless `--resend` is explicitly supplied.
- Each run records row status, timing, error details, screenshots/OCR evidence, and a structured result.

Experimental speed flags reduce evidence or checks. Use them only in a supervised test after understanding the relevant `--help` description; they are not the default safe path.

## Output and audit trail

Runs create a timestamped directory below:

```text
wechat-work-message-validation/logs/<run-id>/
```

The directory contains `run_log.txt`, `result.json`, real-send `evidence_manifest.json`, and enabled screenshots/OCR artifacts. The source workbook receives status, timestamps, check details, batch ID, and duration columns. If the original workbook is locked, the program attempts to preserve results in a recorded copy.

## Tests

Run the pure logic and routing test suite without opening WeCom:

```bash
python -m unittest discover -s tests -v
```

These tests do not replace a supervised GUI smoke test. WeCom UI updates, screen scaling, OS permissions, and OCR runtime changes can affect desktop automation.

## Privacy and responsible use

Only message recipients you are authorized to contact. Do not commit real contact lists, message content, screenshots, logs, generated workbooks, or session evidence. The included example uses fictional recipients and non-deliverable placeholder content.

## Agent usage

`SKILL.md` defines when an Agent should invoke the capability, the preview/confirmation boundary, safety invariants, and completion evidence. The Python program is the current implementation and can later be replaced or complemented by an official API or computer-use tool without changing the Skill's user-level intent.
