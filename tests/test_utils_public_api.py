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
