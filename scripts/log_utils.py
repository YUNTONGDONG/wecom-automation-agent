from __future__ import annotations

import json
from pathlib import Path
import sys
from typing import Any


def log_line(stats: Any, text: str) -> None:
    if stats.log_path is None:
        raise RuntimeError("运行日志路径尚未初始化")
    log_path = stats.log_path
    log_path.parent.mkdir(parents=True, exist_ok=True)
    line = text.rstrip()
    with log_path.open("a", encoding="utf-8") as log:
        log.write(line + "\n")
    if stats.live_log:
        print(line, file=sys.stderr, flush=True)


def save_run_result(stats: Any, payload: dict[str, Any]) -> Path:
    if stats.run_dir is None:
        raise RuntimeError("运行目录尚未初始化")
    result_path = stats.run_dir / "result.json"
    payload["run_dir"] = str(stats.run_dir)
    payload["log_path"] = str(stats.log_path) if stats.log_path else ""
    payload["result_path"] = str(result_path)
    with result_path.open("w", encoding="utf-8") as out:
        json.dump(payload, out, ensure_ascii=False, indent=2)
    return result_path


def row_evidence_summary(item: Any) -> dict[str, Any]:
    return {
        "row": item.row,
        "target": item.target,
        "precheck": item.precheck_status,
        "targetcheck": item.targetcheck_status,
        "postcheck": item.postcheck_status,
        "sent_steps": item.sent_steps,
        "target_screenshot": item.target_screenshot,
        "before_screenshot": item.before_screenshot,
        "after_screenshot": item.after_screenshot,
        "step_events": item.step_events,
        "evidence_files": item.evidence_files,
    }


def save_evidence_manifest(stats: Any, rows: list[Any]) -> Path:
    if stats.run_dir is None:
        raise RuntimeError("运行目录尚未初始化")
    manifest_path = stats.run_dir / "evidence_manifest.json"
    payload = {
        "batch_id": stats.batch_id,
        "started_at": stats.started_at,
        "ended_at": stats.ended_at,
        "rows": [row_evidence_summary(row) for row in rows if row.step_events or row.evidence_files],
    }
    with manifest_path.open("w", encoding="utf-8") as out:
        json.dump(payload, out, ensure_ascii=False, indent=2)
    return manifest_path
