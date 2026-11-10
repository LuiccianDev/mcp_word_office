"""The MCPB manifest must match the registered server surface."""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "manifest.json"


def _manifest() -> dict:
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def test_manifest_tool_list_matches_runtime(registered_tools: dict[str, object]) -> None:
    names = [tool["name"] for tool in _manifest()["tools"]]
    assert len(names) == len(set(names)) == 40
    assert set(names) == set(registered_tools)


def test_manifest_uses_console_entry_point() -> None:
    config = _manifest()["server"]["mcp_config"]
    assert "luiccian-word-mcp" in config["args"]
    assert all("server.py" not in argument for argument in config["args"])


def test_manifest_entry_point_is_module_entry() -> None:
    assert _manifest()["server"]["entry_point"] == "src/word_mcp/__main__.py"


def test_manifest_signature_description_is_accurate() -> None:
    tools = {tool["name"]: tool["description"] for tool in _manifest()["tools"]}
    assert "marker" in tools["add_digital_signature"].lower()
