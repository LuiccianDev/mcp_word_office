"""Version metadata must stay synchronized across every published surface."""

from __future__ import annotations

import importlib.util
import json
import re
import tomllib
from pathlib import Path
from types import ModuleType

from word_mcp import __version__


ROOT = Path(__file__).resolve().parents[1]


def _load_bump_module() -> ModuleType:
    path = ROOT / "scripts" / "bump-version.py"
    spec = importlib.util.spec_from_file_location("bump_version", path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_version_is_synchronized_across_manifests() -> None:
    pyproject = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    docs = (ROOT / "docs" / "TOOLS.md").read_text(encoding="utf-8")
    doc_version = re.search(r"^\*Version: ([\d.]+)\*$", docs, re.MULTILINE)
    assert doc_version
    assert {
        pyproject["project"]["version"],
        manifest["version"],
        doc_version.group(1),
    } == {__version__}


def test_bump_version_updates_every_version_surface(tmp_path: Path) -> None:
    (tmp_path / "pyproject.toml").write_text(
        '[project]\nname = "x"\nversion = "1.0.0"\n', encoding="utf-8"
    )
    (tmp_path / "manifest.json").write_text('{"version": "1.0.0"}\n', encoding="utf-8")
    (tmp_path / "CHANGELOG.md").write_text(
        "# Changelog\n\n## [Unreleased]\n", encoding="utf-8"
    )
    docs = tmp_path / "docs"
    docs.mkdir()
    (docs / "TOOLS.md").write_text(
        "# Tools\n\n*Last updated: January 2026*\n*Version: 1.0.0*\n",
        encoding="utf-8",
    )

    _load_bump_module().bump_files("9.9.9", tmp_path, dry_run=False)

    tools_doc = (docs / "TOOLS.md").read_text(encoding="utf-8")
    assert "*Version: 9.9.9*" in tools_doc
    assert "*Last updated: January 2026*" not in tools_doc
    assert '"version": "9.9.9"' in (tmp_path / "manifest.json").read_text(
        encoding="utf-8"
    )
    assert 'version = "9.9.9"' in (tmp_path / "pyproject.toml").read_text(
        encoding="utf-8"
    )


def test_bump_and_release_workflow_stages_tools_doc() -> None:
    workflow = (ROOT / ".github" / "workflows" / "bump-and-release.yml").read_text(
        encoding="utf-8"
    )
    assert "docs/TOOLS.md" in workflow
