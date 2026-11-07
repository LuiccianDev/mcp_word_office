#!/usr/bin/env python3
"""
Synchronize and update version across pyproject.toml and CHANGELOG.md.

Usage:
    python3 scripts/bump-version.py 2.0.0
    python3 scripts/bump-version.py v2.0.0
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path


def bump_version(new_version: str, repo_root: Path) -> None:
    clean_version = new_version.lstrip("v").strip()
    if not clean_version:
        raise ValueError("Invalid version specified")

    pyproject_path = repo_root / "pyproject.toml"
    changelog_path = repo_root / "CHANGELOG.md"

    # 1. Update pyproject.toml
    if pyproject_path.is_file():
        content = pyproject_path.read_text(encoding="utf-8")
        updated = re.sub(
            r'version\s*=\s*"[^"]+"',
            f'version = "{clean_version}"',
            content,
            count=1,
        )
        pyproject_path.write_text(updated, encoding="utf-8")
        print(f"  [OK] {pyproject_path.name} -> {clean_version}")

    # 2. Update CHANGELOG.md if [Unreleased] is present
    today_str = date.today().isoformat()
    if changelog_path.is_file():
        content = changelog_path.read_text(encoding="utf-8")
        target_header = f"## [{clean_version}] - {today_str}"
        if f"[{clean_version}]" not in content:
            # Promote [Unreleased] to new version
            updated = content.replace(
                "## [Unreleased]",
                f"## [Unreleased]\n\n---\n\n{target_header}",
                1,
            )
            changelog_path.write_text(updated, encoding="utf-8")
            print(f"  [OK] {changelog_path.name} promoted [Unreleased] -> {clean_version}")

    print(f"\nVersion successfully bumped to {clean_version}.")
    print("Next steps:")
    print("  git add pyproject.toml CHANGELOG.md")
    print(f"  git commit -m \"chore: bump version to {clean_version}\"")
    print(f"  git tag -a v{clean_version} -m \"Release v{clean_version}\"")
    print("  git push origin master --tags")


def main() -> None:
    parser = argparse.ArgumentParser(description="Bump project version")
    parser.add_argument("version", help="New version (e.g. 2.0.0 or v2.0.0)")
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parent.parent
    try:
        bump_version(args.version, repo_root)
    except Exception as exc:
        print(f"Error bumping version: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
