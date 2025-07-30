"""Tests for the extended_document_tools module."""

import os
import pytest
import json
from docx import Document
from unittest.mock import patch

# Assuming extended_document_tools.py is in src/word_mcp/tools
from word_mcp.tools import extended_document_tools

@pytest.fixture
def temp_docx_file(tmp_path):
    """Create a temporary .docx file for testing."""
    file_path = tmp_path / "test_document.docx"
    doc = Document()
    doc.add_paragraph("This is the first paragraph.")
    doc.add_paragraph("This is the second paragraph with the word Test.")
    doc.save(file_path)
    return str(file_path)

@pytest.mark.asyncio
async def test_get_paragraph_text_from_document(temp_docx_file):
    """Test getting text from a specific paragraph."""
    result = await extended_document_tools.get_paragraph_text_from_document(temp_docx_file, 1)
    assert result['status'] == "success"
    assert "second paragraph" in result['text']

@pytest.mark.asyncio
async def test_find_text_in_document(temp_docx_file):
    """Test finding text in a document."""
    result = await extended_document_tools.find_text_in_document(temp_docx_file, "Test")
    assert result['status'] == "success"


@pytest.mark.asyncio
@patch('platform.system', return_value='Windows')
@patch('docx2pdf.convert')
async def test_convert_to_pdf_windows(mock_convert, mock_system, temp_docx_file, tmp_path):
    """Test converting a document to PDF on Windows."""
    output_file = tmp_path / "output.pdf"
    result = await extended_document_tools.convert_to_pdf(temp_docx_file, str(output_file))
    mock_convert.assert_called_once_with(temp_docx_file, str(output_file))
    assert result['success']
    assert "successfully converted" in result['message']
