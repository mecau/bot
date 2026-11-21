"""Small, dependency-free helpers shared by MecauAI runtime and tests."""
from __future__ import annotations

import html
import os
import re
from pathlib import Path
from typing import Any

_ALLOWED_TAG_RE = re.compile(r"</?(?:b|strong|i|em|u|s|code|pre|blockquote)>", re.I)
_TAG_TOKEN_RE = re.compile(r"(</?(?:b|strong|i|em|u|s|code|pre|blockquote)>)", re.I)
_ENTITY_OR_CHAR_RE = re.compile(r"&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);|.", re.S)


def _utf16_units(value: str) -> int:
    """Telegram message/entity offsets use UTF-16 code units."""
    return len(str(value).encode("utf-16-le")) // 2


def read_env_int(name: str, default: int, minimum: int | None = None,
                 maximum: int | None = None) -> int:
    """Read an integer environment setting without crashing on a bad value."""
    raw = os.getenv(name)
    try:
        value = int(raw) if raw not in (None, "") else int(default)
    except (TypeError, ValueError):
        value = int(default)
    if minimum is not None:
        value = max(minimum, value)
    if maximum is not None:
        value = min(maximum, value)
    return value


def read_env_float(name: str, default: float, minimum: float | None = None,
                   maximum: float | None = None) -> float:
    """Read a finite float environment setting with safe bounds."""
    import math
    raw = os.getenv(name)
    try:
        value = float(raw) if raw not in (None, "") else float(default)
        if not math.isfinite(value):
            raise ValueError("non-finite")
    except (TypeError, ValueError):
        value = float(default)
    if minimum is not None:
        value = max(minimum, value)
    if maximum is not None:
        value = min(maximum, value)
    return value


def normalize_user_record(record: Any, max_history: int = 6,
                          max_output_chars: int = 4000) -> dict[str, Any]:
    """Repair legacy/malformed user records while preserving unrelated fields."""
    if not isinstance(record, dict):
        record = {}
    result = dict(record)
    mode = result.get("mode")
    result["mode"] = mode if mode in {"ai", "coder", "friend"} else "ai"
    raw_history = result.get("history")
    history = []
    if isinstance(raw_history, list):
        for item in raw_history[-max_history:]:
            if not isinstance(item, dict):
                continue
            role = item.get("role")
            content = item.get("content")
            if role not in {"user", "assistant"} or not isinstance(content, str):
                continue
            history.append({"role": role, "content": content[-8000:]})
    result["history"] = history
    last_output = result.get("last_output")
    result["last_output"] = str(last_output if last_output is not None else "Здесь пока нет ответов.")[-max_output_chars:]
    return result


def normalize_stats_record(record: Any) -> dict[str, Any]:
    """Repair older or partial stats rows before handlers increment counters."""
    result = dict(record) if isinstance(record, dict) else {}
    defaults = {"messages": 0, "voice": 0, "files": 0, "images": 0, "exports": 0}
    for key, default in defaults.items():
        try:
            result[key] = max(0, int(result.get(key, default)))
        except (TypeError, ValueError, OverflowError):
            result[key] = default
    last_action = result.get("last_action")
    result["last_action"] = str(last_action)[:80] if last_action is not None else None
    return result


def safe_filename(value: Any, fallback: str = "file", limit: int = 45) -> str:
    """Return a path-safe readable stem (no path separators or control chars)."""
    text = re.sub(r"[\\/:*?\"<>|\x00-\x1f]+", "_", str(value or ""))
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^A-Za-zА-Яа-яЁё0-9_-]+", "_", text)
    text = text.strip("._-")[:limit].strip("._-")
    return text or fallback


def path_is_within(path: str | os.PathLike[str], root: str | os.PathLike[str]) -> bool:
    """Check containment after resolving symlinks and relative components."""
    try:
        Path(path).resolve().relative_to(Path(root).resolve())
        return True
    except (ValueError, OSError, RuntimeError):
        return False


def format_ai_html(text: Any) -> str:
    """Convert a safe Markdown subset to Telegram HTML; raw HTML is always escaped."""
    raw = str(text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not raw:
        return ""

    placeholders: dict[str, str] = {}

    def hold(fragment: str) -> str:
        key = f"MECAUAIHOLDTOKEN{len(placeholders)}X"
        placeholders[key] = fragment
        return key

    # Protect fenced code before escaping the surrounding text.
    def fenced(match: re.Match[str]) -> str:
        code = match.group(2).strip("\n")
        return hold(f"<pre><code>{html.escape(code, quote=False)}</code></pre>")

    raw = re.sub(r"```([\w.+#-]+)?\s*\n?(.*?)```", fenced, raw, flags=re.S)

    # Protect inline code too, so markdown markers inside code remain literal.
    def inline_code(match: re.Match[str]) -> str:
        return hold(f"<code>{html.escape(match.group(1), quote=False)}</code>")

    raw = re.sub(r"`([^`\n]+)`", inline_code, raw)
    escaped = html.escape(raw, quote=False)

    # Simple headings and lists are readable in Telegram and do not require unsupported tags.
    escaped = re.sub(r"(?m)^\s{0,3}#{1,3}\s+(.+?)\s*$", r"<b>\1</b>", escaped)
    escaped = re.sub(r"(?m)^\s*[-*+]\s+", "• ", escaped)
    escaped = re.sub(r"(?m)^\s*>\s?", "▎ ", escaped)
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escaped, flags=re.S)
    escaped = re.sub(r"__(.+?)__", r"<b>\1</b>", escaped, flags=re.S)
    escaped = re.sub(r"(?<!\*)\*([^*\n]+)\*(?!\*)", r"<i>\1</i>", escaped)
    escaped = re.sub(r"(?<!_)_([^_\n]+)_(?!_)", r"<i>\1</i>", escaped)

    for key, fragment in placeholders.items():
        escaped = escaped.replace(key, fragment)
    return escaped.strip()


def split_telegram_html(text: str, max_chars: int = 3900) -> list[str]:
    """Split Telegram HTML into valid UTF-16-sized chunks, balancing tags/entities."""
    source = str(text or "")
    if max_chars < 100:
        raise ValueError("max_chars must be at least 100")
    if not source:
        return [""]

    tokens = _TAG_TOKEN_RE.split(source)
    chunks: list[str] = []
    current = ""
    visible = 0
    stack: list[str] = []

    def opening_tags() -> str:
        return "".join(f"<{tag}>" for tag in stack)

    def close_tags() -> str:
        return "".join(f"</{tag}>" for tag in reversed(stack))

    def flush() -> None:
        nonlocal current, visible
        if current:
            chunks.append(current + close_tags())
        current = opening_tags()
        visible = 0

    for token in tokens:
        if not token:
            continue
        if token.startswith("<") and token.endswith(">"):
            closing = token.startswith("</")
            tag = token[2:-1].lower() if closing else token[1:-1].lower()
            if closing:
                current += token
                if tag in stack:
                    # Pop the matching tag and any malformed inner tags safely.
                    idx = len(stack) - 1 - stack[::-1].index(tag)
                    stack = stack[:idx]
            else:
                current += token
                stack.append(tag)
            continue

        for atom in _ENTITY_OR_CHAR_RE.findall(token):
            atom_len = _utf16_units(html.unescape(atom))
            if visible and visible + atom_len > max_chars:
                flush()
            current += atom
            visible += atom_len

    if current:
        chunks.append(current + close_tags())
    return chunks or [""]


def split_telegram_text(text: str, max_chars: int = 3900) -> list[str]:
    """Split plain text using Telegram's UTF-16 code-unit length accounting."""
    source = str(text or "")
    if max_chars < 100:
        raise ValueError("max_chars must be at least 100")
    if not source:
        return [""]
    chunks: list[str] = []
    current: list[str] = []
    units = 0
    for char in source:
        char_units = _utf16_units(char)
        if current and units + char_units > max_chars:
            chunks.append("".join(current))
            current = []
            units = 0
        current.append(char)
        units += char_units
    if current:
        chunks.append("".join(current))
    return chunks or [""]


def safe_excel_value(value: Any) -> Any:
    """Prevent untrusted strings from being interpreted as spreadsheet formulas.

    Leading whitespace/control characters are checked too because spreadsheet
    applications may ignore them when deciding whether a cell is a formula.
    """
    if isinstance(value, str):
        probe = value.lstrip(" \t\r\n\ufeff")
        if probe[:1] in {"=", "+", "-", "@"}:
            return "'" + value
    return value


_ALLOWED_EXCEL_FUNCTIONS = {
    "SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "IF", "IFS",
    "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "ROUND", "ROUNDUP",
    "ROUNDDOWN", "ABS", "MEDIAN", "AND", "OR", "NOT", "IFERROR",
}
_FORMULA_CALL_RE = re.compile(r"\b([A-Z][A-Z0-9.]*)\s*\(", re.I)


def safe_excel_formula(value: Any) -> str | None:
    """Allow only simple, non-external formulas using a conservative function allowlist."""
    formula = str(value or "").strip()
    if not formula.startswith("=") or len(formula) > 500:
        return None
    upper = formula.upper()
    if any(token in upper for token in ("!", "[", "]", "HTTP", "WEBSERVICE", "HYPERLINK", "DDE", "CHAR(", "INDIRECT(", "OFFSET(")):
        return None
    functions = set(_FORMULA_CALL_RE.findall(upper))
    if functions and not functions.issubset(_ALLOWED_EXCEL_FUNCTIONS):
        return None
    if not re.fullmatch(r'=[A-Z0-9_.,()+*/%\-\s:$<>="]+', upper):
        return None
    return formula
