# Documentation and Maintainability Audit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Align every repository document with the actual runtime behavior of the 40 MCP tools, and add automated guards that prevent documentation, skills, manifests, and release metadata from drifting again.

**Architecture:** Introduce a machine-readable signature block per tool in `docs/TOOLS.md`, generated and verified by `scripts/sync-tools-reference.py`, which introspects the same functions that `register_tools.py` registers. Add pytest contract tests for tool signatures, documented return shapes, version parity, skill parameters, and packaging. Then correct the remaining prose documents, remove dead utility code, add typing/community metadata, and align pre-commit, CI, and the VS Code workspace.

**Tech Stack:** Python 3.11+, `fastmcp`, `python-docx`, `pytest`, `ruff`, `mypy`, `uv`, Hatchling, GitHub Actions.

**Spec:** Audit findings recorded in this plan (no separate spec file exists). The "Review Focus" section enumerates the failure modes the plan must pin with tests.

## Global Constraints

- All repository documentation, code comments, commit messages, issue templates, and pull requests MUST be written in Technical English.
- User-facing conversation replies remain in Spanish with a formal engineering tone.
- No informal emojis in documentation, comments, release notes, templates, or commits.
- Do not add new runtime dependencies to `pyproject.toml`.
- Preserve exactly 40 canonical tool names. Do not introduce `word_*` aliases.
- `requires-python = ">=3.11"`; supported matrix is 3.11, 3.12, 3.13, 3.14.
- Documentation MUST describe implemented behavior. When a capability is limited, state the limitation instead of omitting it.
- PyPI wheel and sdist scope stays unchanged: `src/word_mcp/`, `README.md`, `LICENSE`, `pyproject.toml`.
- Run every command from the repository root with `uv`.
- Do not bump the package version in this plan. Version synchronization is enforced, not changed.

## Review Focus

1. **Documented parameter that does not exist** (`search_text`, `top_cm`, `editing_mode`): a client copies the example and receives `TypeError` or an ignored argument. Pinned by Task 1 signature parity and Task 2/7 banned-identifier assertions.
2. **Documented return field that the payload does not contain** (`full_text`, `headings`, `word_count`, `next_offset`): downstream logic reads a missing key. Pinned by the per-tool assertions in Task 2.
3. **Release metadata drift** between `pyproject.toml`, `manifest.json`, and `docs/TOOLS.md`: installers and readers disagree about the release. Pinned by Task 3 parity tests plus the `bump-version.py` regression test.
4. **Skill instructions that override model behavior with invalid calls** (`search_text`, `primary_filename`, `protection_type`, `tracked_changes`): the model emits schema-invalid tool calls. Pinned by Task 7.
5. **Removing a helper that is part of the documented `word_mcp.utils` surface**: external imports break. Pinned by the exact `__all__` test in Task 8.

---

## File Structure

**Created**

- `scripts/sync-tools-reference.py` — introspects registered tool functions and rewrites `docs/TOOLS.md` signature blocks.
- `tests/conftest.py` — loads the sync script once per session and exposes `registered_tools`.
- `tests/test_docs_tool_parity.py` — signature and coverage contract.
- `tests/test_docs_content_contract.py` — documented semantics contract.
- `tests/test_version_parity.py` — version surfaces and `bump-version.py` regression.
- `tests/test_skills_contract.py` — skill parameter contract.
- `tests/test_utils_public_api.py` — utility export contract.
- `tests/test_packaging_metadata.py` — `py.typed` and classifier.
- `tests/test_repo_hygiene.py` — community files, pre-commit, CI, workspace.
- `src/word_mcp/py.typed` — PEP 561 marker.
- `CONTRIBUTING.md` — contributor workflow.
- `.gitattributes` — line-ending and binary rules.
- `.github/pull_request_template.md`
- `.github/ISSUE_TEMPLATE/bug_report.yml`
- `.github/ISSUE_TEMPLATE/feature_request.yml`
- `.github/ISSUE_TEMPLATE/config.yml`

**Modified**

- `docs/TOOLS.md` — generated signatures and corrected semantics.
- `README.md`, `SECURITY.md`, `AGENTS.md`, `CHANGELOG.md`, `docs/MCPB.md`.
- `skills/word-document-creator/SKILL.md`, `skills/word-document-creator/reference/elements-cheatsheet.md`, `skills/word-document-creator/reference/presets-and-styling.md`, `skills/word-document-editor/SKILL.md`, `skills/word-document-security/SKILL.md`.
- `manifest.json`, `pyproject.toml`, `.pre-commit-config.yaml`, `word_mcp.code-workspace`, `.github/workflows/ci.yml`, `.github/workflows/codeql.yml`, `scripts/bump-version.py`, `src/word_mcp/server.py`, `src/word_mcp/utils/__init__.py`, `src/word_mcp/utils/document_utils.py`, `src/word_mcp/validation/document_validators.py`, `src/word_mcp/tools/document_tools.py`.

**Deleted**

- `src/word_mcp/utils/extended_document_utils.py`.

---

## Task 1: Tool Signature Contract and Sync Script

**Files:**
- Create: `scripts/sync-tools-reference.py`
- Create: `tests/conftest.py`
- Create: `tests/test_docs_tool_parity.py`
- Modify: `docs/TOOLS.md`

**Interfaces:**
- Consumes: `src/word_mcp/tools/register_tools.py` registration calls in the form `register_read_only(module.function)`, `register_write(module.function)`, `register_destructive(module.function)`.
- Produces: `registered_tools() -> dict[str, Callable[..., Any]]`, `canonical_signature(fn: Callable[..., Any]) -> str`, `render_signature_block(fn: Callable[..., Any]) -> str`, `sync_tools_reference(doc_path: Path = DOC_PATH) -> int`.
- Produces: pytest fixtures `sync_tools` and `registered_tools` in `tests/conftest.py`.

- [ ] **Step 1: Write the failing contract test and fixture**

Create `tests/conftest.py`:

```python
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
```

Create `tests/test_docs_tool_parity.py`:

```python
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_docs_tool_parity.py -v`

Expected: FAIL with `FileNotFoundError` for `scripts/sync-tools-reference.py` because the script does not exist yet.

- [ ] **Step 3: Implement `scripts/sync-tools-reference.py`**

Create the script with this exact behavior:

```python
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
        if not isinstance(func, ast.Attribute) or func.attr not in _REGISTRATION_METHODS:
            continue
        if len(node.args) != 1:
            continue
        target = node.args[0]
        if not isinstance(target, ast.Attribute) or not isinstance(target.value, ast.Name):
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
```

- [ ] **Step 4: Generate the signature blocks**

Run: `uv run python scripts/sync-tools-reference.py`

Expected: `Updated 40 tool signature block(s) in docs/TOOLS.md`

- [ ] **Step 5: Run the tests to verify they pass**

Run: `uv run pytest tests/test_docs_tool_parity.py -v`

Expected: 3 passed.

- [ ] **Step 6: Commit**

```bash
git add scripts/sync-tools-reference.py tests/conftest.py tests/test_docs_tool_parity.py docs/TOOLS.md
git commit -m "test(docs): add tool signature contract and sync script"
```

---

## Task 2: Correct `docs/TOOLS.md` Semantics

**Files:**
- Modify: `docs/TOOLS.md`
- Create: `tests/test_docs_content_contract.py`

**Interfaces:**
- Consumes: signature blocks generated by Task 1.
- Produces: documented behavior that matches the runtime payloads listed in the correction table below.

- [ ] **Step 1: Write the failing content contract test**

Create `tests/test_docs_content_contract.py`:

```python
"""Contract between docs/TOOLS.md prose and runtime behavior."""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOC = ROOT / "docs" / "TOOLS.md"


def section(tool_name: str) -> str:
    text = DOC.read_text(encoding="utf-8")
    match = re.search(
        rf"### `{tool_name}`\n(.*?)(?=\n### `|\n## )", text, re.DOTALL
    )
    assert match, f"missing section for {tool_name}"
    return match.group(1)


def test_build_document_documents_runtime_margins() -> None:
    body = section("build_document")
    for key in ('"top"', '"bottom"', '"left"', '"right"'):
        assert key in body
    assert "top_cm" not in body


def test_build_document_page_numbers_default_false() -> None:
    assert "Default: `False`" in section("build_document")


def test_build_document_documents_space_after_and_orientation() -> None:
    body = section("build_document")
    assert "`space_after`" in body
    assert "`orientation`" in body


def test_build_document_heading_range_and_alignment_default() -> None:
    body = section("build_document")
    assert "1-9" in body
    assert "Default: `left`" in body
    assert "not rendered" in body


def test_get_document_text_documents_status_and_text() -> None:
    body = section("get_document_text")
    assert "`status`" in body
    assert "`text`" in body
    assert "full_text" not in body


def test_get_document_outline_documents_paragraphs_and_tables() -> None:
    body = section("get_document_outline")
    assert "`paragraphs`" in body
    assert "`tables`" in body
    assert "headings" not in body.lower()


def test_get_document_info_documents_properties_and_page_count_limitation() -> None:
    body = section("get_document_info")
    assert "`properties`" in body
    assert "Page count is not computed" in body


def test_list_available_documents_pagination_shape() -> None:
    body = section("list_available_documents")
    assert "next_offset" not in body
    assert "directories_searched" not in body


def test_find_text_in_document_uses_text_to_find() -> None:
    body = section("find_text_in_document")
    assert "`text_to_find`" in body
    assert "`search_results`" in body
    assert "search_text" not in body


def test_footnote_style_parameters_match_runtime() -> None:
    body = section("customize_footnote_style")
    assert "`numbering_format`" in body
    assert "`start_number`" in body
    assert "restart_each_page" not in body


def test_add_footnote_has_no_reference_mark() -> None:
    assert "reference_mark" not in section("add_footnote_to_document")


def test_restricted_editing_parameters_match_runtime() -> None:
    body = section("add_restricted_editing")
    assert "`password`" in body
    assert "`editable_sections`" in body


def test_response_format_scope_is_documented() -> None:
    text = DOC.read_text(encoding="utf-8")
    match = re.search(r"## Response Format\n(.*?)(?=\n## )", text, re.DOTALL)
    assert match
    body = match.group(1)
    assert "does not change the output shape" in body
    assert "get_document_info" in body
    assert "list_available_documents" in body
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_docs_content_contract.py -v`

Expected: multiple FAILs for `top_cm`, `default: True`, `full_text`, `headings`, `page_count`, `next_offset`, `search_text`, `reference_mark`, `restart_each_page`.

- [ ] **Step 3: Apply the correction table to `docs/TOOLS.md`**

| Section | Required documentation |
| :--- | :--- |
| `build_document` | Margins object is `{"top": float, "bottom": float, "left": float, "right": float}` in centimeters. Document `page_numbers` with the exact text `Default: `False``. Document `space_after` (points) and `orientation` (`portrait` or `landscape`). Heading levels 1-9. Paragraph alignment with the exact text `Default: `left``. Callout `title` is accepted but not rendered. |
| `create_from_markdown` | Document that it accepts `markdown_content`, `filename`/`file_path`, `title`, `author`, `preset`, `font_name`, `font_size`, `line_spacing`, `space_after`, `margins`, `header_text`, `footer_text`, `page_numbers`. It has no `orientation` parameter. |
| `batch_add_content` | Returns `status`, `filename`, `total_elements`, `element_counts`, `summary`, `message`. |
| `get_document_info` | Returns `{"status", "properties", "response_format"}`. `properties` contains `author`, `title`, `subject`, `keywords`, `comments`, `category`, `last_modified_by`, `created`, `modified`, `filename`, `path`, `size_kb`. Add the sentence: `Page count is not computed; the OpenXML package does not expose a reliable page total without rendering.` |
| `get_document_text` | Returns `{"status", "text"}`. Remove `full_text`, `paragraph_count`, and `paragraphs`. |
| `get_document_outline` | Returns `{"paragraphs": [{"index", "text", "style"}], "tables": [{"index", "rows", "columns", "preview"}]}`. Paragraph text is truncated to 100 characters plus `...`. Tables include a 3x3 cell preview. There is no heading tree and no `status` key on success. |
| `list_available_documents` | Pagination is `{"page", "page_size", "total", "has_more"}`. Remove `next_offset` and `directories_searched`. State that only top-level `.docx` files in allowed directories are scanned. |
| `merge_documents` | Returns `{"status", "message", "file_path"}`. |
| `add_restricted_editing` | Parameters are `filename`, `password`, `editable_sections`. State that the section list is stored in the sidecar `.protection` file and is advisory. |
| `add_digital_signature` | State that it writes a visible signature block and stores a SHA-256 content hash in the sidecar; it is not a cryptographic signature and does not verify signer identity. |
| `verify_document` | Returns `ProtectionResult`: `status`, `filename`, `protection_type`, `success`, `message`. It reads the sidecar file only. |
| `add_footnote_to_document` | Parameters are `filename`, `paragraph_index`, `footnote_text`. Remove `reference_mark`. |
| `customize_footnote_style` | Parameters are `filename`, `numbering_format` (default `"1, 2, 3"`), `start_number` (default `1`), `font_name`, `font_size`. Remove `restart_each_page`. |
| `get_paragraph_text_from_document` | Returns `{"status", "paragraph": {"index", "text", "style", "is_heading"}}`. Remove `character_count`. |
| `find_text_in_document` | Parameter is `text_to_find`. Returns `{"status", "search_results": {"query", "match_case", "whole_word", "occurrences", "total_count"}}`. Each occurrence has `location`, `position`, `match`, `context`. |
| `convert_to_pdf` | Returns `{"status", "message", "pdf_path"}`. Remove `success`. |
| `add_header`, `add_footer`, `add_hyperlink`, `add_bookmark`, `set_page_layout`, `get_core_properties`, `set_core_properties` | Document `filename` as the primary parameter and `file_path` as an alias. |
| `## Response Format` | Replace the claim that all tools support it with: it is declared by `get_document_info`, `get_document_text`, `get_document_outline`, `list_available_documents`, and `merge_documents`; only `get_document_info` and `list_available_documents` echo the value in the response, and it does not change the output shape. |

Also replace the footer with:

```markdown
*Last updated: November 2026*
*Version: 2.0.1*
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `uv run pytest tests/test_docs_content_contract.py tests/test_docs_tool_parity.py -v`

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add docs/TOOLS.md tests/test_docs_content_contract.py
git commit -m "docs(tools): correct parameters and return shapes"
```

---

## Task 3: Version Parity and Release Automation

**Files:**
- Modify: `scripts/bump-version.py`
- Modify: `docs/TOOLS.md` (footer, if not already correct)
- Create: `tests/test_version_parity.py`

**Interfaces:**
- Consumes: `pyproject.toml` `[project].version`, `manifest.json` `version`, `docs/TOOLS.md` footer, `word_mcp.__version__`.
- Produces: `bump_files(new_version: str, repo_root: Path, dry_run: bool = False) -> None` that also updates `docs/TOOLS.md`.

- [ ] **Step 1: Write the failing parity and bump regression test**

Create `tests/test_version_parity.py`:

```python
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
    (tmp_path / "manifest.json").write_text(
        '{"version": "1.0.0"}\n', encoding="utf-8"
    )
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_version_parity.py -v`

Expected: `test_version_is_synchronized_across_manifests` FAILS because `docs/TOOLS.md` reports `1.0.1`; `test_bump_version_updates_every_version_surface` FAILS because the footer is unchanged.

- [ ] **Step 3: Teach `scripts/bump-version.py` to update the tools footer**

In `bump_files`, after the CHANGELOG block and before the `GITHUB_OUTPUT` block, insert:

```python
    # 4. Update docs/TOOLS.md footer
    tools_doc = repo_root / "docs" / "TOOLS.md"
    if tools_doc.is_file():
        content = tools_doc.read_text(encoding="utf-8")
        updated = re.sub(
            r"(?m)^\*Last updated: .*\*$",
            f"*Last updated: {date.today().strftime('%B %Y')}*",
            content,
            count=1,
        )
        updated = re.sub(
            r"(?m)^\*Version: [^*]+\*$",
            f"*Version: {new_version}*",
            updated,
            count=1,
        )
        if not dry_run:
            tools_doc.write_text(updated, encoding="utf-8")
        print(f"  [OK] docs/TOOLS.md -> {new_version}")
```

Renumber the following comment to `# 5. Export to GITHUB_OUTPUT ...` and update the module docstring line `Synchronizes version across pyproject.toml, manifest.json, and CHANGELOG.md.` to append `, and docs/TOOLS.md`.

- [ ] **Step 4: Run the tests to verify they pass**

Run: `uv run pytest tests/test_version_parity.py -v`

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add scripts/bump-version.py tests/test_version_parity.py docs/TOOLS.md
git commit -m "ci(release): keep documentation version in sync"
```

---

## Task 4: Correct `README.md`

**Files:**
- Modify: `README.md`
- Modify: `tests/test_docs_content_contract.py`

**Interfaces:**
- Consumes: actual `src/word_mcp/` module tree and workflow behavior from `.github/workflows/`.
- Produces: README claims limited to implemented capabilities.

- [ ] **Step 1: Add failing README assertions**

Append to `tests/test_docs_content_contract.py`:

```python
def test_readme_project_structure_lists_actual_modules() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    for module in (
        "footnotes.py",
        "links.py",
        "headers_footers.py",
        "properties.py",
        "footnote_tools.py",
        "property_tools.py",
        "utils/",
        "exception/",
    ):
        assert module in readme


def test_readme_does_not_claim_heading_hierarchy_outline() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "heading hierarchies" not in readme
    assert "paragraph previews and table counts" in readme


def test_readme_does_not_promise_unverified_performance() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "0.25 seconds" not in readme


def test_readme_describes_release_chain_correctly() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "generates GitHub Release, and publishes to PyPI" not in readme
    assert "OOXML hash verification" not in readme


def test_readme_links_tools_reference() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "docs/TOOLS.md" in readme
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_docs_content_contract.py -k readme -v`

Expected: FAIL on missing modules, `heading hierarchies`, `0.25 seconds`, and the release-chain phrase.

- [ ] **Step 3: Apply the README corrections**

1. Replace the "Lightweight Outline Inspection" bullet with:

```markdown
- **Lightweight Outline Inspection**: Extract paragraph previews and table counts without dumping full body text (`get_document_outline`), saving context tokens.
```

2. Replace the "Project Structure" code block with:

```text
src/
└── word_mcp/
    ├── core/
    │   ├── batch.py            # Single-pass builder and markdown parser
    │   ├── content.py          # Headings, paragraphs, tables, pictures
    │   ├── document.py         # Lifecycle, text search, PDF conversion
    │   ├── document_context.py # Read/write document context manager
    │   ├── footnotes.py        # Footnotes and endnotes
    │   ├── formatting.py       # Text and table formatting
    │   ├── headers_footers.py  # Headers, footers, page layout
    │   ├── links.py            # Hyperlinks and bookmarks
    │   ├── properties.py       # Core properties and section layout
    │   ├── protection.py       # Encryption and protection metadata
    │   ├── styles.py           # Style presets and helpers
    │   └── tables.py           # Table helpers
    ├── exception/              # Typed errors and error responses
    ├── models/                 # Pydantic response models
    ├── tools/                  # Registered MCP tools
    │   ├── batch_tools.py
    │   ├── content_tools.py
    │   ├── document_tools.py
    │   ├── footnote_tools.py
    │   ├── format_tools.py
    │   ├── property_tools.py
    │   └── protection_tools.py
    ├── utils/                  # Document and security utilities
    ├── validation/             # Path and file validators
    ├── server.py               # FastMCP server registration
    └── __main__.py             # Console entry point
```

3. Replace the token-optimization row value `**< 0.25 seconds**` with `Single in-memory assembly`.

4. Replace the `bump-and-release.yml` bullet with:

```markdown
- **Automated Version Bump & Tag (`.github/workflows/bump-and-release.yml`)**: 1-click semantic version bump (`patch`, `minor`, `major`). Updates `pyproject.toml`, `manifest.json`, `CHANGELOG.md`, `docs/TOOLS.md`, and `uv.lock`, validates the quality gate, commits, and pushes the `v*` tag. The tag then triggers `release.yml` and, through the published GitHub Release, `publish-pypi.yml`.
```

5. Replace `OOXML hash verification` with `protection metadata and threat model`.

6. Add a reference to the tools reference in the "Key Architectural Highlights" section:

```markdown
- **Complete Tool Reference**: See [`docs/TOOLS.md`](docs/TOOLS.md) for every tool signature, return shape, and limitation.
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `uv run pytest tests/test_docs_content_contract.py -k readme -v`

Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add README.md tests/test_docs_content_contract.py
git commit -m "docs(readme): align structure and CI descriptions"
```

---

## Task 5: Correct `SECURITY.md` and Server Instructions

**Files:**
- Modify: `SECURITY.md`
- Modify: `src/word_mcp/server.py`
- Modify: `tests/test_docs_content_contract.py`

**Interfaces:**
- Consumes: `src/word_mcp/core/protection.py` behavior (`msoffcrypto` encryption, SHA-256 sidecar hash, visible signature block).
- Produces: security documentation that states actual mechanisms and limitations.

- [ ] **Step 1: Add failing security assertions**

Append to `tests/test_docs_content_contract.py`:

```python
def test_security_policy_documents_actual_crypto() -> None:
    policy = (ROOT / "SECURITY.md").read_text(encoding="utf-8")
    assert "msoffcrypto" in policy
    assert "SHA-256" in policy
    assert "SHA-512" not in policy
    assert "100,000" not in policy
    assert "not a cryptographic signature" in policy
    assert "verify_document" in policy
    assert "verify_protection" not in policy


def test_server_instructions_do_not_claim_bidirectional_note_conversion() -> None:
    server = (ROOT / "src" / "word_mcp" / "server.py").read_text(encoding="utf-8")
    assert "and vice versa" not in server
    assert "Endnotes to footnotes" not in server
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_docs_content_contract.py -k "security or server" -v`

Expected: FAIL on `SHA-512`, `100,000`, `verify_protection`, and `and vice versa`.

- [ ] **Step 3: Rewrite the affected SECURITY.md sections**

Replace section `### 2. Password Hashing & OOXML Encryption` with:

```markdown
### 2. Password Protection and Protection Metadata
- `protect_document` encrypts the OOXML container with `msoffcrypto-tool`. The password is not stored in the document or in this repository.
- `unprotect_document` decrypts the container and returns `Incorrect password` when the key is invalid.
- `add_restricted_editing` and `add_digital_signature` write a sidecar JSON file named `<document>.protection` next to the `.docx`. The sidecar stores a SHA-256 hash of the password for restricted-editing verification. This is an unsalted hash, not a password-based key derivation function, and MUST NOT be treated as high-security credential storage.
- `add_digital_signature` writes a visible signature block and stores a SHA-256 hash of the paragraph text. It detects text changes after signing. It is **not a cryptographic signature** and does not verify signer identity.
- `verify_document` reads the sidecar metadata only. It does not validate OOXML package signatures.
```

Add a `### 4. Known Limitations` subsection after the subprocess section:

```markdown
### 4. Known Limitations
- The `.protection` sidecar is advisory metadata. Microsoft Word does not enforce the `editable_sections` list.
- Page count is not computed by `get_document_info`; rendering is required for an authoritative value.
- PDF conversion depends on LibreOffice on Linux and `docx2pdf` on Windows/macOS. Failure is reported as a structured error.
```

Replace the supported-versions table with:

```markdown
| Version | Supported | Notes |
| :--- | :--- | :--- |
| **3.x.x** | :white_check_mark: Yes | Current major branch after the 3.0.0 release |
| **2.x.x** | :x: No | Superseded by 3.0.0; prefixed tool aliases were removed |
| **1.x.x** | :x: No | Deprecated legacy `mcp_word` release |
```

- [ ] **Step 4: Correct the server instructions**

In `src/word_mcp/server.py`, replace:

```python
- **Convert between**: Transform footnotes to endnotes and vice versa
```

with:

```python
- **Convert footnotes to endnotes**: Transform existing footnotes into endnotes (one-way)
```

Replace:

```python
- **Digital signatures**: Sign documents digitally
- **Verification**: Check document protection and signature status
```

with:

```python
- **Signature markers**: Add a visible signature block and a SHA-256 content hash
- **Verification**: Check protection metadata and signature-marker status
```

Replace `- Document protection uses strong encryption` and `- Digital signatures verify document integrity` with:

```python
- Document protection encrypts the OOXML container with `msoffcrypto`
- Signature markers detect text changes; they are not cryptographic signatures
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `uv run pytest tests/test_docs_content_contract.py -k "security or server" -v`

Expected: 2 passed.

- [ ] **Step 6: Commit**

```bash
git add SECURITY.md src/word_mcp/server.py tests/test_docs_content_contract.py
git commit -m "docs(security): describe actual protection mechanisms"
```

---

## Task 6: MCPB Guide and Manifest Entry Point

**Files:**
- Modify: `docs/MCPB.md`
- Modify: `manifest.json`
- Create: `tests/test_manifest_parity.py`

**Interfaces:**
- Consumes: `registered_tools` fixture from Task 1.
- Produces: manifest whose tool list matches the 40 registered tools and whose launch command uses the console script.

- [ ] **Step 1: Write the failing manifest test**

Create `tests/test_manifest_parity.py`:

```python
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_manifest_parity.py -v`

Expected: FAIL because `entry_point` is `src/word_mcp/server.py`, `args` contain `src/word_mcp/server.py`, and the signature description says "Add a digital signature marker to a document" (contains `marker`; this assertion passes already).

- [ ] **Step 3: Update `manifest.json`**

Change:

```json
"entry_point": "src/word_mcp/server.py",
```

to:

```json
"entry_point": "src/word_mcp/__main__.py",
```

Change the `mcp_config.args` array to:

```json
[
  "--directory",
  "${__dirname}",
  "run",
  "luiccian-word-mcp"
]
```

Change the `get_document_outline` description to `"Get paragraph previews and table dimensions from a document."`

- [ ] **Step 4: Rewrite `docs/MCPB.md` in technical English**

Replace the document body with sections that cover, in this order:

1. Overview: MCPB bundles the server manifest for Claude Desktop.
2. Prerequisites: `uv` installed, Python 3.11+, Claude Desktop.
3. Build: `mcpb pack` from the repository root.
4. Install: import the generated `.mcpb` file into Claude Desktop.
5. Configuration: the `directory` user config maps to `MCP_ALLOWED_DIRECTORIES`; select the folder that contains the `.docx` files.
6. Verification: confirm the server starts and that `build_document` is listed.
7. Troubleshooting: `uv` not on `PATH`, invalid directory, missing Python 3.11+.
8. Reference: link to the official MCPB repository.

Do not claim that MCPB removes the need for `uv`; the manifest launches `uv run luiccian-word-mcp`.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `uv run pytest tests/test_manifest_parity.py -v`

Expected: 4 passed.

- [ ] **Step 6: Commit**

```bash
git add manifest.json docs/MCPB.md tests/test_manifest_parity.py
git commit -m "docs(mcpb): rewrite deployment guide and fix manifest entry point"
```

---

## Task 7: Correct the Three Skills

**Files:**
- Modify: `skills/word-document-creator/SKILL.md`
- Modify: `skills/word-document-creator/reference/elements-cheatsheet.md`
- Modify: `skills/word-document-creator/reference/presets-and-styling.md`
- Modify: `skills/word-document-editor/SKILL.md`
- Modify: `skills/word-document-security/SKILL.md`
- Create: `tests/test_skills_contract.py`

**Interfaces:**
- Consumes: `add_hyperlink`, `merge_documents`, `find_text_in_document`, `add_restricted_editing`, `convert_to_pdf` signatures; `render_element` behavior in `src/word_mcp/core/batch.py`.
- Produces: skill examples that use only valid parameters.

- [ ] **Step 1: Write the failing skill contract test**

Create `tests/test_skills_contract.py`:

```python
"""Skill examples must use parameters and behaviors that exist."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILLS = ROOT / "skills"


def text(relative: str) -> str:
    return (SKILLS / relative).read_text(encoding="utf-8")


def test_editor_skill_uses_actual_parameters() -> None:
    body = text("word-document-editor/SKILL.md")
    assert "text_to_find=" in body
    assert "search_text" not in body
    assert "target_filename=" in body
    assert "source_filenames=" in body
    assert "primary_filename" not in body
    assert "secondary_filenames" not in body
    assert '"type": "page_break"' in body
    assert "add_page_break=True" not in body
    assert "add_page_breaks=True" in body
    assert "headings, levels" not in body
    assert "paragraph previews" in body


def test_security_skill_uses_actual_parameters() -> None:
    body = text("word-document-security/SKILL.md")
    assert "editable_sections=" in body
    assert "output_filename=" in body
    assert "protection_type=" not in body
    assert "editing_mode" not in body
    assert "output_pdf_path" not in body
    assert "tracked_changes" not in body
    assert "advisory" in body


def test_creator_skill_argument_hint_lists_all_presets() -> None:
    body = text("word-document-creator/SKILL.md")
    for preset in ("modern", "formal", "academic", "compact"):
        assert preset in body


def test_elements_cheatsheet_matches_renderer() -> None:
    body = text("word-document-creator/reference/elements-cheatsheet.md")
    assert "1 and 9" in body
    assert "Default: `left`" in body
    assert "light background shading" not in body
    assert "always centered" in body
    assert "does not render `title`" in body
    assert "indented italic paragraph" in body


def test_presets_reference_uses_runtime_margin_keys() -> None:
    body = text("word-document-creator/reference/presets-and-styling.md")
    for key in ('"top"', '"bottom"', '"left"', '"right"'):
        assert key in body
    assert "top_cm" not in body
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_skills_contract.py -v`

Expected: FAIL on `search_text`, `primary_filename`, `editing_mode`, `tracked_changes`, heading range, table shading, callout rendering, and `top_cm`.

- [ ] **Step 3: Correct the editor skill**

In `skills/word-document-editor/SKILL.md`:

- Replace `get_document_outline` description with: `Use get_document_outline first to see paragraph previews, paragraph styles, and table dimensions.`
- Rename `search_text=` to `text_to_find=`.
- Replace `primary_filename=` with `target_filename=`.
- Replace `secondary_filenames=` with `source_filenames=`.
- Replace `add_page_break=True` with `add_page_breaks=True`.
- Add a page-break element example in Pattern C:

```json
{ "type": "page_break" }
```

- [ ] **Step 4: Correct the security skill**

In `skills/word-document-security/SKILL.md`:

- Replace the "Password Protection Modes" list (`read_only`, `comments`, `tracked_changes`) with:

```markdown
2. **Protection Modes:**
   - **Container encryption**: `protect_document` encrypts the `.docx` with a password.
   - **Restricted editing metadata**: `add_restricted_editing` records an advisory list of editable sections in the `.protection` sidecar. Microsoft Word does not enforce this list.
   - **Signature marker**: `add_digital_signature` writes a visible signature block and a SHA-256 content hash. It is not a cryptographic signature.
```

- In Pattern A, remove `protection_type="read_only"`.
- Replace Pattern C with:

```python
add_restricted_editing(
    filename="/path/to/borrador_revisores.docx",
    password="SecurePassword2026!",
    editable_sections=["Introduccion", "Conclusiones"]
)
```

- Replace `output_pdf_path=` with `output_filename=`.
- Replace the `verify_document` claim about "cryptographic locks" with: `Run verify_document to inspect sidecar protection metadata before distribution.`

- [ ] **Step 5: Correct the creator references**

In `skills/word-document-creator/reference/elements-cheatsheet.md`:

- Heading: `Integer between 1 and 9 (default: 1).`
- Paragraph: `alignment`: `"left"`, `"center"`, `"right"`, or `"justify"` (Default: `left`). `space_after` inherits the document default (6 pt for the `modern` preset).
- Table: headers are bold only; no background shading is applied. Table alignment is always centered and the `alignment` element key is ignored.
- Callout: the renderer does not render `title`; it renders an indented italic paragraph.

In `skills/word-document-creator/reference/presets-and-styling.md`, replace the custom-margins object with:

```json
"margins": {
  "top": 2.5,
  "bottom": 2.5,
  "left": 3.0,
  "right": 2.5
}
```

- [ ] **Step 6: Run the tests to verify they pass**

Run: `uv run pytest tests/test_skills_contract.py -v`

Expected: 5 passed.

- [ ] **Step 7: Commit**

```bash
git add skills tests/test_skills_contract.py
git commit -m "docs(skills): correct tool parameters and rendering behavior"
```

---

## Task 8: Remove Dead Utilities and Stale Comments

**Files:**
- Delete: `src/word_mcp/utils/extended_document_utils.py`
- Modify: `src/word_mcp/utils/__init__.py`
- Modify: `src/word_mcp/utils/document_utils.py`
- Modify: `src/word_mcp/validation/document_validators.py`
- Modify: `src/word_mcp/tools/document_tools.py`
- Create: `tests/test_utils_public_api.py`

**Interfaces:**
- Consumes: `word_mcp.utils` public exports.
- Produces: `word_mcp.utils.__all__ == {"create_document_copy", "ensure_docx_extension", "extract_document_text", "get_document_structure"}`.

- [ ] **Step 1: Write the failing public API test**

Create `tests/test_utils_public_api.py`:

```python
"""The word_mcp.utils public surface must be explicit and free of dead code."""

from __future__ import annotations

import importlib
import importlib.util


def test_utils_exports_match_expected_surface() -> None:
    utils = importlib.import_module("word_mcp.utils")
    assert set(utils.__all__) == {
        "create_document_copy",
        "ensure_docx_extension",
        "extract_document_text",
        "get_document_structure",
    }
    for name in utils.__all__:
        assert hasattr(utils, name)


def test_unused_helpers_are_removed() -> None:
    utils = importlib.import_module("word_mcp.utils")
    for name in (
        "get_document_properties",
        "find_paragraph_by_text",
        "find_and_replace_text",
        "get_paragraph_text",
        "find_text",
    ):
        assert not hasattr(utils, name)


def test_extended_document_utils_module_is_removed() -> None:
    assert importlib.util.find_spec("word_mcp.utils.extended_document_utils") is None
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_utils_public_api.py -v`

Expected: FAIL because the removed names are still exported.

- [ ] **Step 3: Delete the dead module and trim the exports**

1. Delete `src/word_mcp/utils/extended_document_utils.py`.
2. Replace `src/word_mcp/utils/__init__.py` with:

```python
"""
Utility functions for the Word Document Server.
This package contains utility modules for file operations and document handling.
"""

from word_mcp.utils.document_utils import (
    create_document_copy,
    ensure_docx_extension,
    extract_document_text,
    get_document_structure,
)


__all__ = [
    "create_document_copy",
    "ensure_docx_extension",
    "extract_document_text",
    "get_document_structure",
]
```

3. In `src/word_mcp/utils/document_utils.py`, delete `get_document_properties`, `find_paragraph_by_text`, and `find_and_replace_text`. Remove the now-unused imports `Any`, `CT_CoreProperties`, and `validate_docx_file` only if no remaining function uses them; `extract_document_text` and `get_document_structure` still use `validate_docx_file`, and `Any` is still used by `get_document_structure`.
4. Update the module docstring from `including extracting properties, text, and structure from .docx files.` to `including extracting text, structure, and copying documents.`
5. Remove the comment `#! Observe that this function is not used in the code, it is left for future use` above `ensure_docx_extension`.

- [ ] **Step 4: Fix stale module and inline comments**

1. Replace the module docstring of `src/word_mcp/validation/document_validators.py` with:

```python
"""Validation decorators for Word document tool parameters.

This module provides decorators and helpers that verify file existence,
extension, readability, and write access before a tool executes.
"""
```

2. In `src/word_mcp/tools/document_tools.py`, replace the comment block inside `get_document_info`:

```python
        # Properties usually doesn't need full doc open via DocumentContext 
        # but for consistency we could, however document_utils.get_document_properties 
        # might be optimized. Let's stick to core if we had one for properties.
        # We have core_get_core_properties but it takes doc.
```

with:

```python
        # Core properties are read from the open document context.
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `uv run pytest tests/test_utils_public_api.py -v`

Expected: 3 passed.

Run: `rg "extended_document_utils|find_paragraph_by_text|find_and_replace_text|get_document_properties" src tests`

Expected: no matches.

- [ ] **Step 6: Commit**

```bash
git add -A src/word_mcp/utils src/word_mcp/validation/document_validators.py src/word_mcp/tools/document_tools.py tests/test_utils_public_api.py
git commit -m "refactor(utils): remove unused document helpers"
```

---

## Task 9: Ship the PEP 561 Marker

**Files:**
- Create: `src/word_mcp/py.typed`
- Modify: `pyproject.toml`
- Create: `tests/test_packaging_metadata.py`

**Interfaces:**
- Consumes: Hatchling wheel target `packages = ["src/word_mcp"]`.
- Produces: wheel that contains `word_mcp/py.typed` and metadata with `Typing :: Typed`.

- [ ] **Step 1: Write the failing packaging test**

Create `tests/test_packaging_metadata.py`:

```python
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_packaging_metadata.py -v`

Expected: FAIL because `py.typed` is missing and the classifier is absent.

- [ ] **Step 3: Add the marker and classifier**

Create an empty file `src/word_mcp/py.typed`.

Add `"Typing :: Typed",` to the `classifiers` list in `pyproject.toml`, after the Python version classifiers.

- [ ] **Step 4: Run the tests and verify the wheel contents**

Run: `uv run pytest tests/test_packaging_metadata.py -v`

Expected: 2 passed.

Run:

```bash
uv build
uv run python -c "import glob, zipfile; wheel = sorted(glob.glob('dist/*.whl'))[-1]; print([n for n in zipfile.ZipFile(wheel).namelist() if n.endswith('py.typed')])"
```

Expected: `['word_mcp/py.typed']`

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/py.typed pyproject.toml tests/test_packaging_metadata.py
git commit -m "build: ship py.typed marker"
```

---

## Task 10: Add Contributor and GitHub Community Files

**Files:**
- Create: `CONTRIBUTING.md`
- Create: `.gitattributes`
- Create: `.github/pull_request_template.md`
- Create: `.github/ISSUE_TEMPLATE/bug_report.yml`
- Create: `.github/ISSUE_TEMPLATE/feature_request.yml`
- Create: `.github/ISSUE_TEMPLATE/config.yml`
- Create: `tests/test_repo_hygiene.py`

**Interfaces:**
- Consumes: quality-gate commands already used by CI.
- Produces: contributor documentation and repository metadata referenced by `AGENTS.md`.

- [ ] **Step 1: Write the failing hygiene test**

Create `tests/test_repo_hygiene.py`:

```python
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_repo_hygiene.py -v`

Expected: FAIL because the files do not exist.

- [ ] **Step 3: Create `CONTRIBUTING.md`**

Include these sections and commands:

- Development setup: `git clone`, `uv sync --all-groups`.
- Quality gate: `uv run ruff check src tests scripts`, `uv run ruff format --check src tests scripts`, `uv run mypy src`, `uv run pytest`.
- Documentation contract: after changing a tool signature, run `uv run python scripts/sync-tools-reference.py`, then `uv run pytest tests/test_docs_tool_parity.py`.
- Commit conventions: Conventional Commits in English; use `[skip ci]` only for documentation-only commits.
- Changelog: add user-visible changes under `## [Unreleased]` using `Added`, `Changed`, `Deprecated`, `Removed`, `Fixed`, `Security`.
- Release process: use `.github/workflows/bump-and-release.yml`; do not tag or publish manually.
- Security reports: follow `SECURITY.md`; do not open public issues.

- [ ] **Step 4: Create `.gitattributes`**

```gitattributes
* text=auto eol=lf

*.docx binary
*.pdf binary
*.png binary
*.jpg binary
*.jpeg binary
*.gif binary
*.whl binary
*.tar.gz binary
```

- [ ] **Step 5: Create the pull request template**

Use these sections: Summary, Motivation, Changes, Verification, Documentation, Checklist. The checklist must include:

```markdown
- [ ] `uv run ruff check src tests scripts`
- [ ] `uv run mypy src`
- [ ] `uv run pytest`
- [ ] `uv run python scripts/sync-tools-reference.py` (when tool signatures changed)
- [ ] `CHANGELOG.md` updated under `[Unreleased]`
```

- [ ] **Step 6: Create the issue templates**

`bug_report.yml` fields: `name: Bug report`, `description`, `labels: ["bug"]`, and a `body` with required inputs for `word-mcp version`, Python version, operating system, steps to reproduce, expected behavior, actual behavior, and logs.

`feature_request.yml` fields: `name: Feature request`, `description`, `labels: ["enhancement"]`, and a `body` with required inputs for problem statement, proposed solution, alternatives considered, and affected tool names.

`config.yml`:

```yaml
blank_issues_enabled: false
contact_links:
  - name: Security vulnerability
    url: https://github.com/LuiccianDev/word-mcp/security/policy
    about: Report security issues privately; do not open a public issue.
```

- [ ] **Step 7: Run the tests to verify they pass**

Run: `uv run pytest tests/test_repo_hygiene.py -v`

Expected: 3 passed.

- [ ] **Step 8: Commit**

```bash
git add CONTRIBUTING.md .gitattributes .github tests/test_repo_hygiene.py
git commit -m "docs: add contributor and GitHub community files"
```

---

## Task 11: Align Pre-commit, CI, and Workspace Configuration

**Files:**
- Modify: `.pre-commit-config.yaml`
- Modify: `word_mcp.code-workspace`
- Modify: `.github/workflows/ci.yml`
- Modify: `.github/workflows/codeql.yml`
- Modify: `tests/test_repo_hygiene.py`

**Interfaces:**
- Consumes: quality gate commands `uv run ruff check src tests scripts`, `uv run mypy src`, `uv run pytest`.
- Produces: identical scope across pre-commit, CI, and the workspace.

- [ ] **Step 1: Add failing configuration assertions**

Append to `tests/test_repo_hygiene.py`:

```python
def test_pre_commit_covers_source_tests_and_scripts() -> None:
    body = (ROOT / ".pre-commit-config.yaml").read_text(encoding="utf-8")
    assert "^(src|tests|scripts)/.*\\.py$" in body


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
        body = (ROOT / ".github" / "workflows" / workflow).read_text(
            encoding="utf-8"
        )
        assert ".gitattributes" in body
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_repo_hygiene.py -v`

Expected: FAIL for pre-commit scope, workspace exclusions, and `.gitattributes` path filtering.

- [ ] **Step 3: Update `.pre-commit-config.yaml`**

Change both Ruff hooks from `files: ^src/.*\.py$` to `files: ^(src|tests|scripts)/.*\.py$`.

Keep the mypy hook scoped to `files: ^src/.*\.py$`, because `pyproject.toml` excludes `tests/` from mypy.

- [ ] **Step 4: Update `word_mcp.code-workspace`**

Remove the entries `"src": true`, `"tests": true`, `"docs": false`, `"skills": true`, `"dist": true`, and `"build": true` from both `files.exclude` and `search.exclude`.

Add `"scripts/"` to the `Ruff Check` and `Ruff Format` argument arrays.

Change the `Quality Gate` command to:

```text
uv run ruff check src/ tests/ scripts/ && uv run mypy src/ && uv run pytest -v
```

- [ ] **Step 5: Add `.gitattributes` to workflow path filters**

In both `.github/workflows/ci.yml` and `.github/workflows/codeql.yml`, add `- '.gitattributes'` to each `paths-ignore` list.

- [ ] **Step 6: Run the tests to verify they pass**

Run: `uv run pytest tests/test_repo_hygiene.py -v`

Expected: 7 passed.

- [ ] **Step 7: Commit**

```bash
git add .pre-commit-config.yaml word_mcp.code-workspace .github/workflows/ci.yml .github/workflows/codeql.yml tests/test_repo_hygiene.py
git commit -m "chore: align pre-commit, CI, and workspace configuration"
```

---

## Task 12: Changelog and Full Verification

**Files:**
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: all prior tasks.
- Produces: release notes and an evidence-backed quality gate run.

- [ ] **Step 1: Add the changelog entries**

Under `## [Unreleased]`, add:

```markdown
### Added
- Documentation contract tests for tool signatures, return shapes, skills, version parity, packaging, and repository hygiene.
- `scripts/sync-tools-reference.py` to regenerate `docs/TOOLS.md` signature blocks from the registered tools.
- `src/word_mcp/py.typed` and the `Typing :: Typed` classifier.
- `CONTRIBUTING.md`, GitHub issue and pull request templates, and `.gitattributes`.

### Changed
- Aligned `docs/TOOLS.md`, `README.md`, `SECURITY.md`, `docs/MCPB.md`, and the three skills with the implemented runtime behavior.
- `scripts/bump-version.py` now updates the `docs/TOOLS.md` version footer.
- Pre-commit, CI, and the VS Code workspace now share the same `src`, `tests`, and `scripts` scope.

### Removed
- Deleted the unused `word_mcp.utils.extended_document_utils` module and the unused `get_document_properties`, `find_paragraph_by_text`, and `find_and_replace_text` helpers.

### Fixed
- `manifest.json` now launches the `luiccian-word-mcp` console entry point instead of `src/word_mcp/server.py`.
- Corrected stale module docstrings in the validators package and the `ensure_docx_extension` comment.
```

- [ ] **Step 2: Run the full quality gate**

Run: `uv run ruff check src tests scripts`

Expected: `All checks passed!`

Run: `uv run ruff format --check src tests scripts`

Expected: `N files already formatted`

Run: `uv run mypy src`

Expected: `Success: no issues found in 39 source files` (40 before Task 8 deletes `extended_document_utils.py`).

Run: `uv run pytest`

Expected: all tests pass, including the 72 pre-existing tests plus the new contract tests.

Run: `uv lock --check`

Expected: exit code 0 and no lockfile changes.

- [ ] **Step 3: Build and inspect the distribution**

Run:

```bash
uv build
uv run python -c "import glob, zipfile; wheel = sorted(glob.glob('dist/*.whl'))[-1]; names = zipfile.ZipFile(wheel).namelist(); assert 'word_mcp/py.typed' in names; assert not any(n.startswith(('tests/', 'scripts/', 'docs/', 'skills/')) for n in names); print('wheel scope OK')"
```

Expected: `wheel scope OK`

- [ ] **Step 4: Review the diff for scope violations**

Run: `git status --short` and `git diff --stat`

Expected: only the files listed in this plan are modified, created, or deleted. No dependency changes, no version bump, no tool renames.

- [ ] **Step 5: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs(changelog): record documentation and maintenance work"
```

---

## Self-Review Notes

- Every audit finding maps to a task: `docs/TOOLS.md` (Tasks 1-2), `README.md` (Task 4), `AGENTS.md`/`SECURITY.md` (Tasks 5 and 10), `docs/MCPB.md` and `manifest.json` (Task 6), skills (Task 7), dead code and stale comments (Task 8), `py.typed` (Task 9), community files (Task 10), configuration drift (Task 11), release automation (Task 3), changelog and verification (Task 12).
- No task changes a tool name, adds a runtime dependency, or alters the PyPI distribution boundary.
- The signature block format is the single source of truth for parameter names, defaults, and annotations; prose is only allowed to describe semantics and limitations.
- `bump-version.py` is covered by a regression test so the footer cannot drift on the next release.
