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


