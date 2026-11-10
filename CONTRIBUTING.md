# Contributing to Word MCP

Thank you for your interest in contributing to **Word MCP Server** (`luiccian-word-mcp`)! This document outlines our development workflows, quality gates, and standards.

---

## 1. Development Setup

We manage dependencies and virtual environments with Astral `uv`:

```bash
# Clone the repository
git clone https://github.com/LuiccianDev/word-mcp.git
cd word-mcp

# Sync environment with all dependencies and groups
uv sync --all-groups
```

---

## 2. Quality Gate

Before submitting a pull request, all checks must pass locally:

```bash
# 1. Lint and format checks
uv run ruff check src tests scripts
uv run ruff format --check src tests scripts

# 2. Static type analysis
uv run mypy src

# 3. Automated test suite
uv run pytest
```

---

## 3. Documentation Contract

When modifying or adding tool signatures in `src/word_mcp/tools/`:

1. Regenerate signature blocks in `docs/TOOLS.md`:
   ```bash
   uv run python scripts/sync-tools-reference.py
   ```
2. Verify documentation parity:
   ```bash
   uv run pytest tests/test_docs_tool_parity.py
   ```

---

## 4. Commit Conventions

- Write all commit messages in **Technical English** following the [Conventional Commits](https://www.conventionalcommits.org/) specification (e.g., `feat(core): ...`, `fix(tools): ...`, `docs(mcpb): ...`).
- Use `[skip ci]` or `[ci skip]` in the commit message **only** for commits that touch documentation, markdown, or git metadata exclusively.

---

## 5. Changelog Maintenance

Record user-facing improvements, fixes, or breaking changes in [`CHANGELOG.md`](CHANGELOG.md) under the `## [Unreleased]` section following [Keep a Changelog](https://keepachangelog.com/):
- `Added`: for new features or capabilities.
- `Changed`: for changes in existing functionality.
- `Deprecated`: for soon-to-be-removed features.
- `Removed`: for now-removed features.
- `Fixed`: for any bug fixes.
- `Security`: in case of vulnerabilities.

---

## 6. Release Process

Do **not** perform manual git tagging or PyPI uploading. Releases are automated via the GitHub Actions workflow `.github/workflows/bump-and-release.yml`:
1. The workflow triggers `scripts/bump-version.py`.
2. It promotes changelog entries, updates version surfaces, tags `v*`, and triggers automated deployment to PyPI using Trusted Publishing (OIDC).

---

## 7. Security Reports

Do **not** open public GitHub issues for security vulnerabilities. Follow the responsible disclosure instructions in [`SECURITY.md`](SECURITY.md) by contacting the maintainer directly.
