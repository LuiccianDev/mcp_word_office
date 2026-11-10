"""Repository hygiene files must exist and document the real quality gate."""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_community_health_files_exist() -> None:
    for relative in (
        "CONTRIBUTING.md",
        ".gitattributes",
        ".github/pull_request_template.md",
        ".github/ISSUE_TEMPLATE/bug_report.yml",
        ".github/ISSUE_TEMPLATE/feature_request.yml",
        ".github/ISSUE_TEMPLATE/config.yml",
    ):
        assert (ROOT / relative).is_file(), relative


def test_contributing_documents_quality_gate() -> None:
    body = (ROOT / "CONTRIBUTING.md").read_text(encoding="utf-8")
    for command in (
        "uv sync --all-groups",
        "uv run pytest",
        "uv run ruff check src tests scripts",
        "uv run mypy src",
        "uv run python scripts/sync-tools-reference.py",
    ):
        assert command in body


def test_agents_referenced_metadata_file_exists() -> None:
    agents = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
    assert ".gitattributes" in agents
    assert (ROOT / ".gitattributes").is_file()
