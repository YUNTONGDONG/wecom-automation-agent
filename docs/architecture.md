# Architecture

## Agent layer

```text
Natural-language request
    |
    v
Responses API or deterministic Mock model
    |
    v
Bounded function-calling loop
    |
    v
Strict SendPlan schema + workspace path policy
    |
    v
Preview tool -> SQLite task + workbook snapshot -> human approval
    |
    v
Plan-hash verification -> simulation
```

The Agent and automation engine are separate trust domains. The model may propose tool arguments, but application code validates every field. The model never receives the approval secret or approval token and cannot expose arbitrary shell commands. In the MVP, the Agent toolbox intentionally excludes both simulation and real sending; simulation is a separate human-operated CLI transition.

SQLite persists tasks, token hashes, execution idempotency keys, per-delivery keys, and append-only audit events. Approval is valid for one exact task and plan hash, expires after ten minutes, and is atomically consumed once. Before execution, the system re-hashes every workbook plus the selected row values. A changed file invalidates approval.

An exclusive lock file protects the single desktop/WeCom session. The lock contains a random ownership token and can only be removed by its owner. Database uniqueness constraints provide a second layer against duplicate execution or message delivery.

Task states follow:

```text
DRAFT -> PREVIEWED -> AWAITING_CONFIRMATION -> AUTHORIZED
      -> SIMULATING -> COMPLETED / FAILED
```

Future real execution must use a separate `EXECUTING` transition and retain the existing GUI safety checks. It must not reuse simulation as an implicit authorization for sending.

## Execution flow

```text
Excel workbook
    |
    v
Workbook parser and eligibility checks
    |
    v
SendRow execution plan -----> preview JSON / terminal table
    |
    v (only with --execute --yes)
Schedule gate and foreground-window guard
    |
    v
Target search -> result match -> conversation-title verification
    |
    v
Clipboard write -> input verification -> one send action
    |
    v
Newest-message and failure-state verification
    |
    +-----> Excel status/timing/check write-back
    +-----> result.json
    +-----> evidence_manifest.json + screenshots/OCR
```

## Components

- `send_from_excel_1v1_text.py` owns orchestration, platform GUI adapters, scheduling, evidence capture, and status transitions.
- `excel_utils.py` normalizes headers and validates workbook files.
- `log_utils.py` creates per-run directories and writes concise and structured logs.
- `macos_vision_ocr.py` provides the macOS Vision OCR adapter.
- `ocr_windows_image.ps1` provides the Windows OCR adapter.
- `SKILL.md` defines the Agent-level invocation contract and safety boundaries independently from the current implementation.

## State model

A row begins as eligible or skipped. During real execution it moves through target confirmation, input confirmation, send action, and post-send confirmation. A confirmed outgoing bubble becomes `已发送` (or `超时已发送`). An action with ambiguous post-send evidence becomes `发送待核对`, preventing a blind retry. A pre-send validation failure becomes `发送失败` or `已跳过` without clicking Send.

## Trust boundaries

The workbook, clipboard, OCR output, GUI focus, and visible message history are all treated as fallible inputs. No single OCR result proves delivery. Success requires agreement between the intended row, the opened target, the composed content, the newly visible outgoing state, and the recorded execution event.

## Incremental modularization

Keep the verified GUI sequence stable while extracting pure logic first: schema normalization, schedule parsing, target matching, message planning, and state transitions. Add adapter interfaces only after characterization tests capture current macOS and Windows behavior. This limits regressions in the most environment-sensitive code.
