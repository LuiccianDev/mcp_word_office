"""Tests for the document_tools module."""

import os
import pytest
from docx import Document

# Assuming document_tools.py is in src/word_mcp/tools
from word_mcp.tools import document_tools

@pytest.fixture
def temp_docx_file(tmp_path):
    """Create a temporary .docx file for testing."""
    file_path = tmp_path / "test_document.docx"
    doc = Document()
    doc.add_paragraph("This is a test document.")
    doc.save(file_path)
    # Set allowed directory for testing
    os.environ['MCP_ALLOWED_DIRECTORIES'] = str(tmp_path)
    return str(file_path)

@pytest.mark.asyncio
async def test_create_document(tmp_path):
    """Test creating a new document."""
    os.environ['MCP_ALLOWED_DIRECTORIES'] = str(tmp_path)
    file_path = tmp_path / "new_document.docx"
    result = await document_tools.create_document(str(file_path), title="Test Title", author="Test Author")
    assert result['status'] == 'success'
    assert isinstance(result, dict)
    assert os.path.exists(file_path)

@pytest.mark.asyncio
async def test_get_document_info(temp_docx_file):
    """Test getting document information."""
    result = await document_tools.get_document_info(temp_docx_file)
    assert isinstance(result, dict)
    assert 'title' in result

@pytest.mark.asyncio
async def test_get_document_text(temp_docx_file):
    """Test getting document text."""
    result = await document_tools.get_document_text(temp_docx_file)
    assert "This is a test document." in result

@pytest.mark.asyncio
async def test_get_document_outline(temp_docx_file):
    """Test getting document outline."""
    result = await document_tools.get_document_outline(temp_docx_file)
    assert isinstance(result, dict)

@pytest.mark.asyncio
async def test_list_available_documents(temp_docx_file, tmp_path):
    """Test listing available documents."""
    result = await document_tools.list_available_documents(str(tmp_path))
    assert result['status'] == 'success'
    assert len(result['documents']) > 0

@pytest.mark.asyncio
async def test_copy_document(temp_docx_file, tmp_path):
    """Test copying a document."""
    dest_path = tmp_path / "copied_document.docx"
    result = await document_tools.copy_document(temp_docx_file, str(dest_path))
    assert result['status'] == 'success'
    assert os.path.exists(dest_path)


