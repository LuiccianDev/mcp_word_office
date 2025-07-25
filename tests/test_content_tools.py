"""Tests for content_tools module using pytest.

This module contains unit tests for the content manipulation functions in content_tools.py.
"""
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pytest
import pytest_asyncio
from pathlib import Path
from tempfile import NamedTemporaryFile
from unittest.mock import MagicMock, patch, AsyncMock

from docx import Document
from docx.document import Document as DocumentType
from docx.table import Table

from mcp_word_server.tools.content_tools import (
    add_heading,
    add_paragraph,
    add_table,
    add_picture,
    add_page_break,
    add_table_of_contents,
    delete_paragraph,
    search_and_replace,
    _validate_heading_level,
    _extract_headings,
    _populate_table
)

# Add parent directory to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


@pytest_asyncio.fixture
async def temp_docx():
    """Create a temporary docx file for testing."""
    with NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
        temp_path = temp_file.name
    
    # Create a basic document for testing
    doc = Document()
    doc.save(temp_path)
    
    yield temp_path
    
    # Cleanup
    if os.path.exists(temp_path):
        os.unlink(temp_path)


@pytest.fixture
def mock_document():
    """Create a mock document for testing."""
    doc = MagicMock(spec=DocumentType)
    doc.styles = {'Normal': MagicMock()}
    return doc

@pytest.mark.parametrize("level,expected", [
    (1, 1),
    (5, 5),
    (9, 9)
])
def test_validate_heading_level_valid(level, expected):
    """Test _validate_heading_level with valid levels."""
    assert _validate_heading_level(level) == expected


@pytest.mark.parametrize("invalid_level", [0, 10, -1])
def test_validate_heading_level_invalid(invalid_level):
    """Test _validate_heading_level with invalid levels."""
    with pytest.raises(ValueError):
        _validate_heading_level(invalid_level)
    
@pytest.mark.asyncio
async def test_add_heading_success(temp_docx):
    """Test adding a heading successfully."""
    # Setup mock
    mock_doc = MagicMock(spec=DocumentType)
    mock_doc.styles = {'Normal': MagicMock(), 'Heading 1': MagicMock()}
    
    with patch('mcp_word_server.tools.content_tools.Document', return_value=mock_doc):
        # Test
        result = await add_heading(temp_docx, "Test Heading", 1)
        
        # Assertions
        assert "Test Heading" in result
        assert "added to" in result
        mock_doc.add_heading.assert_called_once_with("Test Heading", level=1)
    
@pytest.mark.asyncio
async def test_add_paragraph_success(temp_docx):
    """Test adding a paragraph successfully."""
    # Setup mock
    mock_doc = MagicMock(spec=DocumentType)
    mock_doc.styles = {'Normal': MagicMock()}
    
    with patch('mcp_word_server.tools.content_tools.Document', return_value=mock_doc):
        # Test with default style
        result = await add_paragraph(temp_docx, "Test paragraph")
        
        # Assertions
        assert "Paragraph added to" in result
        mock_doc.add_paragraph.assert_called_once_with("Test paragraph")
    
@pytest.mark.asyncio
async def test_add_table_success(temp_docx):
    """Test adding a table successfully."""
    # Setup mock
    mock_doc = MagicMock(spec=DocumentType)
    mock_table = MagicMock(spec=Table)
    mock_doc.add_table.return_value = mock_table
    
    with patch('mcp_word_server.tools.content_tools.Document', return_value=mock_doc):
        # Test with data
        test_data = [["A1", "B1"], ["A2", "B2"]]
        result = await add_table(temp_docx, 2, 2, test_data)
        
        # Assertions
        assert "Table (2x2) added to" in result
        mock_doc.add_table.assert_called_once_with(rows=2, cols=2)
    
@pytest.mark.asyncio
async def test_add_picture_success(temp_docx, tmp_path):
    """Test adding a picture successfully."""
    # Create a temporary image file
    img_path = tmp_path / "test.png"
    img_path.write_bytes(b'dummy image data')
    
    # Setup mock document
    mock_doc = MagicMock(spec=DocumentType)
    
    with patch('mcp_word_server.tools.content_tools.Document', return_value=mock_doc), \
         patch('os.path.exists', return_value=True), \
         patch('os.path.isfile', return_value=True), \
         patch('os.access', return_value=True):
        
        # Test
        result = await add_picture(temp_docx, str(img_path), width=6.0)
        
        # Assertions
        assert "added to" in result
        mock_doc.add_picture.assert_called_once()
    
@pytest.mark.asyncio
async def test_add_page_break_success(temp_docx):
    """Test adding a page break successfully."""
    # Setup mock
    mock_doc = MagicMock(spec=DocumentType)
    
    with patch('mcp_word_server.tools.content_tools.Document', return_value=mock_doc):
        # Test
        result = await add_page_break(temp_docx)
        
        # Assertions
        assert "Page break added to" in result
        mock_doc.add_page_break.assert_called_once()
    
def test_populate_table():
    """Test populating a table with data."""
    # Create a mock table
    mock_table = MagicMock(spec=Table)
    
    # Set up cell access
    mock_table.cell = MagicMock()
    mock_table.cell.return_value.text = ""
    
    # Test data
    test_data = [
        ["A1", "B1"],
        ["A2", "B2"]
    ]
    
    # Call the function
    _populate_table(mock_table, test_data, 2, 2)
    
    # Verify the table was populated correctly
    assert mock_table.cell.call_count == 4  # 2x2 = 4 cells
    
@pytest.mark.asyncio
async def test_search_and_replace(temp_docx):
    """Test search and replace functionality."""
    # Setup mock document
    mock_doc = MagicMock(spec=DocumentType)

    with patch('mcp_word_server.tools.content_tools.Document', return_value=mock_doc), \
         patch('mcp_word_server.tools.content_tools.find_and_replace_text', return_value=1) as mock_replace:
        # Test
        result = await search_and_replace(temp_docx, "Old", "New")

        # Assertions
        assert "Replaced 1 occurrence(s) of 'Old' with 'New' in" in result
        mock_replace.assert_called_once_with(mock_doc, "Old", "New")
    
def test_extract_headings():
    """Test extracting headings from a document."""
    # Create a mock document with headings
    mock_doc = MagicMock()
    mock_paragraph = MagicMock()
    mock_paragraph.style.name = 'Heading 1'
    mock_paragraph.text = "Test Heading"
    mock_doc.paragraphs = [mock_paragraph]
    
    # Test
    headings = _extract_headings(mock_doc, 3)
    
    # Assertions
    assert len(headings) == 1
    assert headings[0]['text'] == "Test Heading"
    assert headings[0]['level'] == 1
