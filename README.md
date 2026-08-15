# WeCom Automation Agent

An LLM-powered, safety-gated Agent that turns natural-language requests into validated WeCom task plans and invokes a deterministic Excel/GUI automation engine. It combines Responses API function calling, structured validation, human approval, target verification, post-action checks, and evidence-backed audit logs.

> This project controls a real messaging client. Start with preview mode and a supervised 1-3 row test. Never use production contact data in a public repository.

## Why this is more than a send script

The repository now contains three explicit layers:

```text
LLM Agent Runtime
  -> Structured plan + bounded function-calling loop
  -> Human approval bound to the exact plan hash
  -> Whitelisted preview/simulation tools
  -> Deterministic GUI automation engine
```

The GUI execution pipeline treats every action as untrusted until it is verified:

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
├── pyproject.toml
├── .gitignore
├── agents/openai.yaml              # Codex UI metadata
├── docs/architecture.md
├── evals/
│   ├── cases.jsonl                 # Versioned planning and safety dataset
│   └── run_evals.py                # Mock or live-model evaluation runner
├── examples/example_tasks.xlsx     # Synthetic, non-production input
├── src/wecom_agent/
│   ├── agent.py                    # Bounded Agent tool loop
│   ├── model_client.py             # OpenAI Responses + local Mock clients
│   ├── schemas.py                  # Structured plan validation
│   ├── permissions.py              # Plan-hash approval gate
│   ├── state_store.py              # Persistent task state machine
│   ├── tools.py                    # Whitelisted Agent tools
│   └── cli.py
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

Create a virtual environment and install the project:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -e .
```

On Windows activate with `.venv\Scripts\activate`.

## Agent MVP quick start

The default CLI uses a deterministic local Mock client. It demonstrates the complete Agent control flow without an API key and cannot send a message.

Preview a natural-language request:

```bash
wecom-agent --workspace . preview "预览 examples/example_tasks.xlsx"
```

The response includes a generated `task_id`. Approve that exact previewed plan. The returned approval token expires after ten minutes and is not stored in plaintext:

```bash
wecom-agent --workspace . approve <task_id>
```

Run the safe simulation:

```bash
wecom-agent --workspace . simulate <task_id> --approval-token <token>
```

Simulation never opens WeCom and never launches the sender subprocess. Task, approval, execution, delivery, and audit state is stored in ignored `.agent-state/agent.db`. Inspect a task without exposing its approval token:

```bash
wecom-agent --workspace . status <task_id>
```

To use the live model for planning, set `OPENAI_API_KEY` and add `--live` to the preview command. The default model is `gpt-5.6-terra`, overridable through `--model` or `AGENT_MODEL`.

```bash
wecom-agent --workspace . preview --live "预览 examples/example_tasks.xlsx"
```

The live model can only call registered tools. Arbitrary shell execution and real message delivery are not exposed in this MVP.

## Repeatable model evaluations

The versioned evaluation set measures planning behavior without dispatching any tool or opening WeCom. It covers valid plans, row selection, scheduling, missing information, invalid paths, unsafe requests, and prompt-injection attempts.

Run the deterministic baseline:

```bash
PYTHONPATH=src python3 evals/run_evals.py --provider mock
```

Run the same cases against a real model after setting `OPENAI_API_KEY`:

```bash
PYTHONPATH=src python3 evals/run_evals.py \
  --provider openai --model gpt-5.6-terra --repetitions 3 \
  --baseline-output evals/baselines/gpt-5.6-terra.json
```

Reports are written below ignored `evals/results/` and contain per-case outputs, latency, task/action accuracy, argument accuracy, safe-behavior rate, unsafe-tool-call rate, and safe-downgrade rate. A dangerous send request may safely become a preview because preview cannot deliver a message; unknown tools, multiple tool calls, and policy violations still fail. Pin the model name and repetition count when comparing runs. A non-zero exit code means the configured quality thresholds failed, so the runner can also be used as a CI quality gate. Live-model output can still vary; repetitions expose that variance while the dataset and scoring code remain fixed.

The optional baseline file contains only aggregate metrics, quality thresholds, and hashes of the dataset and system prompt. It excludes prompts, tool arguments, errors, and model text, so it can be reviewed before committing. The default gate requires at least 95% pass rate, 95% argument accuracy, and 100% safe handling of unsafe cases.

## Automation engine quick start

Preview the synthetic example without sending anything:

```bash
python3 scripts/send_from_excel_1v1_text.py \
  --folder . \
  --workbook examples/example_tasks.xlsx \
  --json
```

For a real workbook, review preview output first. Then run the confirmed batch:

```bash
python3 scripts/send_from_excel_1v1_text.py \
  --folder /absolute/project/folder \
  --workbook /absolute/path/tasks.xlsx \
  --execute --yes --json
```

Use selected rows:

```bash
python3 scripts/send_from_excel_1v1_text.py \
  --folder . --workbook /absolute/path/tasks.xlsx \
  --row 2 --row 5 --execute --yes --json
```

Respect planned send times:

```bash
python3 scripts/send_from_excel_1v1_text.py \
  --folder . --workbook /absolute/path/tasks.xlsx \
  --respect-send-time --execute --yes --json
```

Run `python3 scripts/send_from_excel_1v1_text.py --help` for advanced execution, evidence, row-range, and diagnostic flags.

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

## Agent safety model

- Model output is parsed into a strict `SendPlan` and validated again in application code.
- Workbook paths must resolve inside the configured workspace.
- The Agent has a bounded number of tool rounds and rejects unknown tool names.
- Approval tokens are cryptographically bound to the canonical plan hash.
- Approval tokens expire after ten minutes, are stored only as hashes, and can be consumed once.
- Any target, row, schedule, workbook, resend, file content, size, or modification-time change invalidates approval.
- Model-generated approval tokens cannot pass verification.
- SQLite uniqueness constraints reject repeated execution and delivery idempotency keys.
- An exclusive workspace lock prevents two desktop execution processes from running together.
- The MVP exposes preview and simulation only; real GUI execution is deliberately absent from the Agent toolbox.

## GUI safety model

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

Run the Agent, permission, schema, pure logic, and routing test suites without opening WeCom:

```bash
python3 -m unittest discover -s tests -v
```

These tests do not replace a supervised GUI smoke test. WeCom UI updates, screen scaling, OS permissions, and OCR runtime changes can affect desktop automation.

Pull requests and pushes to `main` run the same unit tests plus the deterministic Mock evaluation on Python 3.10, 3.11, and 3.12 through GitHub Actions. CI does not receive an OpenAI API key, call a live model, open WeCom, or send messages.

## Privacy and responsible use

Only message recipients you are authorized to contact. Do not commit real contact lists, message content, screenshots, logs, generated workbooks, or session evidence. The included example uses fictional recipients and non-deliverable placeholder content.

## Agent usage

`SKILL.md` defines when an Agent should invoke the capability, the preview/confirmation boundary, safety invariants, and completion evidence. The Python program is the current implementation and can later be replaced or complemented by an official API or computer-use tool without changing the Skill's user-level intent.
