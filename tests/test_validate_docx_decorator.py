"""Tests for the validate_docx_file decorator in file_utils.py."""
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pytest
import asyncio
from pathlib import Path
from typing import Dict, Any

# Import the decorator and related functions
from mcp_word_server.validation.document_validators import validate_docx_file

# Test data directory path
TEST_DATA_DIR = Path(__file__).parent.parent / "test_data"

# Create a sample .docx file for testing
SAMPLE_DOCX = TEST_DATA_DIR / "sample.docx"
INVALID_DOCX = TEST_DATA_DIR / "invalid.docx"

@pytest.fixture(scope="module", autouse=True)
def setup_test_files():
    """Setup test files before tests and clean up afterward."""
    # Create test_data directory if it doesn't exist
    TEST_DATA_DIR.mkdir(exist_ok=True)
    
    # Create a sample .docx file
    from docx import Document
    doc = Document()
    doc.add_paragraph("Test document for unit testing")
    doc.save(SAMPLE_DOCX)
    
    # Create an invalid .docx file (empty file)
    with open(INVALID_DOCX, "wb") as f:
        f.write(b"This is not a valid docx file")
    
    yield  # This is where the testing happens
    
    # Cleanup
    if SAMPLE_DOCX.exists():
        SAMPLE_DOCX.unlink()
    if INVALID_DOCX.exists():
        INVALID_DOCX.unlink()
    if TEST_DATA_DIR.exists() and not any(TEST_DATA_DIR.iterdir()):
        TEST_DATA_DIR.rmdir()

def test_validate_docx_file_sync():
    """Test the decorator with a synchronous function."""
    @validate_docx_file("doc_path")
    def process_document(doc_path: str) -> Dict[str, str]:
        return {"status": "success", "path": doc_path}
    
    # Test with valid docx
    result = process_document(doc_path=str(SAMPLE_DOCX))
    assert result["status"] == "success"
    assert str(SAMPLE_DOCX) in result["path"]
    
    # Test with non-existent file
    result = process_document(doc_path="nonexistent.docx")
    assert "error" in result
    assert "does not exist" in result["error"]
    
    # Test with invalid docx
    result = process_document(doc_path=str(INVALID_DOCX))
    assert "error" in result
    assert "not a valid Word document" in result["error"]

@pytest.mark.asyncio
async def test_validate_docx_file_async():
    """Test the decorator with an asynchronous function."""
    @validate_docx_file("doc_path")
    async def async_process_document(doc_path: str) -> Dict[str, str]:
        await asyncio.sleep(0.01)  # Simulate async operation
        return {"status": "async_success", "path": doc_path}
    
    # Test with valid docx
    result = await async_process_document(doc_path=str(SAMPLE_DOCX))
    assert result["status"] == "async_success"
    assert str(SAMPLE_DOCX) in result["path"]
    
    # Test with non-existent file
    result = await async_process_document(doc_path="nonexistent_async.docx")
    assert "error" in result
    assert "does not exist" in result["error"]

def test_validate_docx_file_positional_args():
    """Test the decorator with positional arguments."""
    @validate_docx_file("doc_path")
    def process_document(doc_path: str, extra_param: str) -> Dict[str, str]:
        return {"status": "success", "path": doc_path, "extra": extra_param}
    
    result = process_document(str(SAMPLE_DOCX), "test")
    assert result["status"] == "success"
    assert result["extra"] == "test"

def test_validate_docx_file_missing_param():
    """Test the decorator when the required parameter is missing."""
    @validate_docx_file("nonexistent_param")
    def process_document(doc_path: str) -> Dict[str, str]:
        return {"status": "success"}
    
    result = process_document(doc_path=str(SAMPLE_DOCX))
    assert "error" in result
    assert "Parameter 'nonexistent_param' not found" in result["error"]

def test_validate_docx_file_non_docx():
    """Test the decorator with a non-.docx file."""
    test_file = TEST_DATA_DIR / "test.txt"
    try:
        # Create a non-docx file
        with open(test_file, "w") as f:
            f.write("This is a text file, not a docx")
        
        @validate_docx_file("doc_path")
        def process_document(doc_path: str) -> Dict[str, str]:
            return {"status": "success"}
        
        result = process_document(doc_path=str(test_file))
        assert "error" in result
        assert "not a .docx document" in result["error"]
    finally:
        if test_file.exists():
            test_file.unlink()
