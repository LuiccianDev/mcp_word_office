# Bump Version and Release Patch v2.0.2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Execute the automated GitHub Actions release pipeline to bump the version to v2.0.2 (patch), tag the release, generate PEP 740 attestation, create the GitHub Release, and publish `luiccian-word-mcp` v2.0.2 to PyPI.

**Architecture:** Utilize repository automation defined in `.github/workflows/bump-and-release.yml`, which invokes `scripts/bump-version.py --type patch`. Downstream push triggers `release.yml` to build artifacts and create the GitHub Release with extracted changelog notes, which in turn triggers `publish-pypi.yml` using Trusted Publishing (OIDC).

**Tech Stack:** GitHub Actions, GitHub CLI (`gh`), Astral `uv`, Python 3.11+, PyPI Trusted Publishing.

**Spec:** Repository release specifications documented in [AGENTS.md](file:///C:/Users/DAVID/Desktop/WorksSpacesCoding/word-mcp/AGENTS.md) (Section 1 and Section 6).

## Global Constraints

- Adhere strictly to Semantic Versioning (SemVer 2.0.0) and Keep a Changelog.
- Do not perform manual version tagging or ad-hoc PyPI publishing; execution must run through `.github/workflows/bump-and-release.yml`.
- All commit messages and release notes must be in Technical English with zero informal emojis.
- `master` must remain clean, tested, and passing all quality gates (`ruff`, `mypy`, `pytest`).

## Review Focus

- Ensure `CHANGELOG.md` has an active `[Unreleased]` section with the security fixes before bumping.
- Validate that the calculated next version is strictly `2.0.2` (patch increment from `2.0.1`).
- Verify that `manifest.json`, `pyproject.toml`, and `docs/TOOLS.md` stay synchronized.
- Ensure the git tag `v2.0.2` does not already exist before triggering.
- Confirm PyPI package `luiccian-word-mcp` receives the wheel and sdist with provenance attestation.

---

### Task 1: Pre-Release Quality Gates and Dry-Run Verification

**Files:**
- Test: `tests/`
- Script: `scripts/bump-version.py`
- Modify: None (read-only verification)

**Interfaces:**
- Consumes: `scripts/bump-version.py --type patch --dry-run`
- Produces: Confidence in local state and clean quality gates before triggering release workflow.

- [ ] **Step 1: Run complete test suite**

Run: `uv run pytest`
Expected: 125 passed.

- [ ] **Step 2: Run linter and formatting check**

Run: `uv run ruff check src tests scripts && uv run ruff format --check .`
Expected: 0 errors, all files formatted.

- [ ] **Step 3: Run static type checking**

Run: `uv run mypy src`
Expected: Success: no issues found in 36 source files.

- [ ] **Step 4: Execute dry-run of bump script**

Run: `uv run python scripts/bump-version.py --type patch --dry-run`
Expected: Output showing next version `2.0.2` (Tag: `v2.0.2`) and promotions across `pyproject.toml`, `manifest.json`, `CHANGELOG.md`, and `docs/TOOLS.md`.

---

### Task 2: Trigger and Monitor Automated GitHub Release Pipeline

**Files:**
- Workflow: `.github/workflows/bump-and-release.yml`
- Workflow: `.github/workflows/release.yml`
- Workflow: `.github/workflows/publish-pypi.yml`

**Interfaces:**
- Consumes: `gh workflow run bump-and-release.yml -f bump_type=patch`
- Produces: Tag `v2.0.2` on `origin/master`, GitHub Release `v2.0.2`, PyPI package `luiccian-word-mcp==2.0.2`.

- [ ] **Step 1: Ensure local working tree is clean and synced with remote master**

Run: `git status && git pull origin master`
Expected: Working tree clean, up to date with `origin/master`.

- [ ] **Step 2: Dispatch `bump-and-release.yml` workflow**

Run: `gh workflow run bump-and-release.yml -f bump_type=patch`
Expected: Workflow dispatch acknowledged by GitHub API.

- [ ] **Step 3: Monitor `bump-and-release.yml` run to completion**

Run: `gh run list --workflow=bump-and-release.yml --limit 1` and `gh run watch <run-id>`
Expected: Completed with status `success`.

- [ ] **Step 4: Monitor downstream `Release - Build & Publish Artifacts` workflow**

Run: `gh run list --workflow=release.yml --limit 1` and `gh run watch <run-id>`
Expected: Completed with status `success`, creating GitHub release `v2.0.2`.

- [ ] **Step 5: Monitor downstream `Publish to PyPI` workflow**

Run: `gh run list --workflow=publish-pypi.yml --limit 1` and `gh run watch <run-id>`
Expected: Completed with status `success`, deploying to PyPI with PEP 740 attestation.

---

### Task 3: Local Synchronization and Post-Release Verification

**Files:**
- Local Git repository (`origin/master` and tags)

**Interfaces:**
- Consumes: Tag `v2.0.2` and commit on `origin/master`
- Produces: Fully synchronized local clone and verification of PyPI availability.

- [ ] **Step 1: Fetch and pull remote release commit and tags**

Run: `git pull origin master --tags`
Expected: Fast-forward merge of release commit and fetch of tag `v2.0.2`.

- [ ] **Step 2: Verify GitHub Release notes and assets**

Run: `gh release view v2.0.2`
Expected: View showing version 2.0.2 title, security changelog entries, and attached `.whl` and `.tar.gz` assets.

- [ ] **Step 3: Verify PyPI publication status**

Run: `uv pip install --dry-run luiccian-word-mcp==2.0.2` or check via curl
Expected: Version 2.0.2 resolves successfully from PyPI index.
