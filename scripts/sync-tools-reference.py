#!/usr/bin/env python3
"""Synchronize docs/TOOLS.md signature blocks with runtime tool signatures."""

from __future__ import annotations

import ast
import importlib
import inspect
import re
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
DOC_PATH = ROOT / "docs" / "TOOLS.md"
REGISTER_TOOLS = SRC / "word_mcp" / "tools" / "register_tools.py"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

_REGISTRATION_METHODS = {
    "register_read_only",
    "register_write",
    "register_destructive",
}


def registered_tools() -> dict[str, Callable[..., Any]]:
    """Return tool functions in registration order, keyed by canonical name."""
    tree = ast.parse(REGISTER_TOOLS.read_text(encoding="utf-8"))
    tools: dict[str, Callable[..., Any]] = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        method = (
            func.attr if isinstance(func, ast.Attribute) else getattr(func, "id", None)
        )
        if method not in _REGISTRATION_METHODS:
            continue
        if len(node.args) != 1:
            continue
        target = node.args[0]
        if not isinstance(target, ast.Attribute) or not isinstance(
            target.value, ast.Name
        ):
            continue
        module = importlib.import_module(f"word_mcp.tools.{target.value.id}")
        tools[target.attr] = getattr(module, target.attr)
    return tools


def canonical_signature(fn: Callable[..., Any]) -> str:
    """Render a stable single-line signature for documentation."""
    return f"{fn.__name__}{inspect.signature(fn, eval_str=True)}"


def render_signature_block(fn: Callable[..., Any]) -> str:
    return f"```python\n{canonical_signature(fn)}\n```"


def sync_tools_reference(doc_path: Path = DOC_PATH) -> int:
    """Replace the signature block under each documented tool heading."""
    text = doc_path.read_text(encoding="utf-8")
    updated = 0
    for name, fn in registered_tools().items():
        pattern = re.compile(
            rf"(^### `{re.escape(name)}`\n\n)(?:```python\n.*?\n```\n\n)?",
            re.MULTILINE | re.DOTALL,
        )
        text, count = pattern.subn(
            lambda match: f"{match.group(1)}{render_signature_block(fn)}\n\n",
            text,
            count=1,
        )
        updated += count
    if updated:
        doc_path.write_text(text, encoding="utf-8", newline="\n")
    return updated


def main() -> int:
    updated = sync_tools_reference()
    print(f"Updated {updated} tool signature block(s) in {DOC_PATH.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
