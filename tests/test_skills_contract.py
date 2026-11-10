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


def test_editor_tools_reference_matches_runtime() -> None:
    body = text("word-document-editor/reference/tools-reference.md")
    assert "headings tree" not in body.lower()
    assert "word count estimate" not in body.lower()
    assert "paragraph previews" in body
