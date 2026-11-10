"""Contract between docs/TOOLS.md and the registered tool functions."""

from __future__ import annotations

from pathlib import Path
from types import ModuleType


ROOT = Path(__file__).resolve().parents[1]
DOC = ROOT / "docs" / "TOOLS.md"


def _signature_block(text: str, tool_name: str) -> str | None:
    marker = f"### `{tool_name}`\n\n```python\n"
    start = text.find(marker)
    if start == -1:
        return None
    start += len(marker)
    end = text.find("\n```", start)
    return text[start:end] if end != -1 else None


def test_registers_exactly_forty_tools(registered_tools: dict[str, object]) -> None:
    assert len(registered_tools) == 40


def test_every_registered_tool_has_signature_block(
    registered_tools: dict[str, object],
) -> None:
    text = DOC.read_text(encoding="utf-8")
    missing = [
        name for name in registered_tools if _signature_block(text, name) is None
    ]
    assert missing == []


def test_documented_signatures_match_runtime(
    sync_tools: ModuleType, registered_tools: dict[str, object]
) -> None:
    text = DOC.read_text(encoding="utf-8")
    mismatched = {
        name: (sync_tools.canonical_signature(fn), _signature_block(text, name))
        for name, fn in registered_tools.items()
        if _signature_block(text, name) != sync_tools.canonical_signature(fn)
    }
    assert mismatched == {}
