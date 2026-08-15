from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Iterable
from zipfile import is_zipfile


PREFERRED_NAME_KEYWORDS = ("企业微信", "群发", "codex", "1v1")


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def split_paths(raw: str) -> list[str]:
    text = cell_text(raw)
    if not text:
        return []
    for sep in ("\r\n", "\n", "\r", "；", ";", "，", ","):
        text = text.replace(sep, "|")
    return [part.strip().strip('"') for part in text.split("|") if part.strip()]


def first_col(headers: dict[str, int], names: Iterable[str]) -> int | None:
    for name in names:
        if name in headers:
            return headers[name]
    return None


def col_value(ws: Any, row: int, headers: dict[str, int], names: Iterable[str]) -> str:
    col = first_col(headers, names)
    return cell_text(ws.cell(row, col).value) if col else ""


def ensure_columns(ws: Any, headers: dict[str, int], names: Iterable[str]) -> dict[str, int]:
    for name in names:
        if name not in headers:
            col = ws.max_column + 1
            ws.cell(1, col).value = name
            headers[name] = col
    return headers


def parse_parts(message_type: str, message: str, image_paths: list[str], document_paths: list[str]) -> dict[str, bool]:
    text = cell_text(message_type)
    if not text:
        return {
            "text": bool(message),
            "image": bool(image_paths),
            "file": bool(document_paths),
        }
    return {
        "text": "文字" in text or "文本" in text or "消息" in text,
        "image": "图片" in text or "图" in text,
        "file": "文件" in text or "文档" in text or "附件" in text,
    }


def find_workbook(folder: Path) -> Path:
    candidates = [
        path
        for path in folder.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm"}
        and not path.name.startswith(".~")
        and is_zipfile(path)
    ]
    if not candidates:
        raise FileNotFoundError(f"No .xlsx/.xlsm workbook found in {folder}")

    def score(path: Path) -> tuple[int, float]:
        name = path.name.lower()
        keyword_score = sum(1 for key in PREFERRED_NAME_KEYWORDS if key.lower() in name)
        return keyword_score, path.stat().st_mtime

    return sorted(candidates, key=score, reverse=True)[0]


def resolve_paths(folder: Path, raw: str) -> list[Path]:
    paths: list[Path] = []
    for item in split_paths(raw):
        path = Path(item).expanduser()
        if not path.is_absolute():
            path = folder / path
        paths.append(path)
    return paths


def worksheet_headers(ws: Any, row: int = 1) -> dict[str, int]:
    return {
        str(ws.cell(row, col).value).strip(): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(row, col).value is not None and str(ws.cell(row, col).value).strip()
    }
