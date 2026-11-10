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


def test_workspace_does_not_hide_source_folders() -> None:
    workspace = json.loads(
        (ROOT / "word_mcp.code-workspace").read_text(encoding="utf-8")
    )
    for key in ("files.exclude", "search.exclude"):
        exclude = workspace["settings"][key]
        for folder in ("src", "tests", "skills"):
            assert exclude.get(folder) is not True, (key, folder)


def test_workspace_quality_gate_includes_scripts() -> None:
    workspace = json.loads(
        (ROOT / "word_mcp.code-workspace").read_text(encoding="utf-8")
    )
    tasks = {task["label"]: task for task in workspace["tasks"]["tasks"]}
    assert "scripts/" in tasks["Quality Gate"]["command"]
    assert "scripts/" in tasks["Ruff Check"]["args"]
    assert "scripts/" in tasks["Ruff Format"]["args"]


def test_ci_ignores_gitattributes_only_changes() -> None:
    for workflow in ("ci.yml", "codeql.yml"):
        body = (ROOT / ".github" / "workflows" / workflow).read_text(encoding="utf-8")
        assert ".gitattributes" in body
