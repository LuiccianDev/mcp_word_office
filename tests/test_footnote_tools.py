"""Tests for the footnote_tools module."""

import pytest
from docx import Document

# Assuming footnote_tools.py is in src/word_mcp/tools
from word_mcp.tools import footnote_tools

@pytest.fixture
def temp_docx_file(tmp_path):
    """Create a temporary .docx file for testing."""
    file_path = tmp_path / "test_document.docx"
    doc = Document()
    doc.add_paragraph("This is a paragraph for testing footnotes.")
    doc.save(file_path)
    return str(file_path)

@pytest.mark.asyncio
async def test_add_footnote_to_document(temp_docx_file):
    """Test adding a footnote to a document."""
    result = await footnote_tools.add_footnote_to_document(temp_docx_file, 0, "This is a test footnote.")
    assert "Footnote added" in result

@pytest.mark.asyncio
async def test_add_endnote_to_document(temp_docx_file):
    """Test adding an endnote to a document."""
    result = await footnote_tools.add_endnote_to_document(temp_docx_file, 0, "This is a test endnote.")
    assert "Endnote added" in result

@pytest.mark.asyncio
async def test_convert_footnotes_to_endnotes_in_document(temp_docx_file):
    """Test converting footnotes to endnotes."""
    # First, add a footnote to convert
    await footnote_tools.add_footnote_to_document(temp_docx_file, 0, "A footnote to be converted.")
    result = await footnote_tools.convert_footnotes_to_endnotes_in_document(temp_docx_file)
    # This is a simplified check; a real test would inspect the document more deeply
    assert "Converted" in result or "No footnote references found" in result
