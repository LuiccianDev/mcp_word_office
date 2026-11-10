"""Packaging metadata must advertise the typed package."""

from __future__ import annotations

import tomllib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_package_ships_py_typed_marker() -> None:
    assert (ROOT / "src" / "word_mcp" / "py.typed").is_file()


def test_project_declares_typed_classifier() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert "Typing :: Typed" in data["project"]["classifiers"]
