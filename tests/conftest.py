"""Shared fixtures for repository contract tests."""

from __future__ import annotations

import importlib.util
from pathlib import Path
from types import ModuleType

import pytest


ROOT = Path(__file__).resolve().parents[1]


def _load_sync_module() -> ModuleType:
    path = ROOT / "scripts" / "sync-tools-reference.py"
    spec = importlib.util.spec_from_file_location("sync_tools_reference", path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="session")
def sync_tools() -> ModuleType:
    return _load_sync_module()


@pytest.fixture(scope="session")
def registered_tools(sync_tools: ModuleType) -> dict[str, object]:
    return sync_tools.registered_tools()
