# Architecture

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
