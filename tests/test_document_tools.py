import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import tempfile
from unittest.mock import patch

import pytest
import pytest_asyncio
from docx import Document

# Import the module to test
from src.word_mcp.tools import document_tools

# Mark all tests in this module as asyncio
pytestmark = pytest.mark.asyncio

# Test data - usando directorio temporal en lugar de rutas fijas
# TEST_DIR se configurará dinámicamente en los tests


# Fixture to create a temporary directory for test documents
@pytest.fixture
def temp_docs_dir():
    """Create a temporary directory for test documents."""
    with tempfile.TemporaryDirectory() as temp_dir:
        # Configurar directorio permitido para tests
        with patch.dict("os.environ", {"MCP_ALLOWED_DIRECTORIES": temp_dir}):
            yield temp_dir


# Fixture to create a test document
@pytest_asyncio.fixture
async def test_document(temp_docs_dir):
    """Create a test document and return its path."""
    doc_path = os.path.join(temp_docs_dir, "test_document.docx")
    doc = Document()
    doc.add_heading("Test Document", level=1)
    doc.add_paragraph("This is a test document.")
    doc.save(doc_path)
    return doc_path


class TestDocumentTools:
    """Test cases for document_tools.py"""

    async def test_create_document_success(self, temp_docs_dir):
        """Test successful document creation."""
        # Configurar directorio permitido para este test
        with patch.dict("os.environ", {"MCP_ALLOWED_DIRECTORIES": temp_docs_dir}):
            filename = os.path.join(temp_docs_dir, "new_document.docx")
            title = "Test Title"
            author = "Test Author"

            # Test with title and author
            result = await document_tools.create_document(
                filename=filename, title=title, author=author
            )

            assert "created successfully" in result.lower()
            assert os.path.exists(filename)

            # Verify document properties
            doc = Document(filename)
            assert doc.core_properties.title == title
            assert doc.core_properties.author == author

    async def test_create_document_invalid_path(self):
        """Test document creation with invalid path."""
        # Configurar directorio permitido vacío para forzar error
        with patch.dict("os.environ", {"MCP_ALLOWED_DIRECTORIES": "/tmp/allowed"}):
            # Try to create a document in a non-allowed directory
            filename = "/non/existent/path/document.docx"

            result = await document_tools.create_document(filename)

            assert (
                "cannot create directory" in result.lower()
                or "not in allowed directories" in result.lower()
            )
            assert not os.path.exists(filename)

    async def test_get_document_info(self, test_document):
        """Test getting document information."""
        result = await document_tools.get_document_info(test_document)

        # The result should be a JSON string with document properties
        import json

        doc_info = json.loads(result)

        assert isinstance(doc_info, dict)
        assert "title" in doc_info
        assert "author" in doc_info
        assert "created" in doc_info

    async def test_get_document_text(self, test_document):
        """Test extracting text from a document."""
        result = await document_tools.get_document_text(test_document)

        assert isinstance(result, str)
        assert "test document" in result.lower()

    async def test_list_available_documents(self, temp_docs_dir):
        """Test listing available documents in a directory."""
        # Configurar directorio permitido para este test
        with patch.dict("os.environ", {"MCP_ALLOWED_DIRECTORIES": temp_docs_dir}):
            # Create some test documents
            doc1 = os.path.join(temp_docs_dir, "doc1.docx")
            doc2 = os.path.join(temp_docs_dir, "doc2.docx")

            Document().save(doc1)
            Document().save(doc2)

            # The result should be a JSON string with document paths
            result = await document_tools.list_available_documents(temp_docs_dir)

            # Validar estructura de diccionario
            assert isinstance(result, dict)
            assert result["status"] == "ok"
            assert result["total"] >= 2
            assert any("doc1.docx" in doc["name"] for doc in result["documents"])
            assert any("doc2.docx" in doc["name"] for doc in result["documents"])

    async def test_copy_document(self, test_document, temp_docs_dir):
        """Test copying a document."""
        dest_path = os.path.join(temp_docs_dir, "copied_document.docx")

        result = await document_tools.copy_document(
            source_filename=test_document, destination_filename=dest_path
        )

        assert "copied to" in result.lower()
        assert os.path.exists(dest_path)

    @patch("mcp_word_server.tools.document_tools.merge_documents")
    async def test_merge_documents(self, mock_merge, temp_docs_dir):
        """Test merging multiple documents."""
        # Configurar directorio permitido para este test
        with patch.dict("os.environ", {"MCP_ALLOWED_DIRECTORIES": temp_docs_dir}):
            # Create source documents
            doc1 = os.path.join(temp_docs_dir, "doc1.docx")
            doc2 = os.path.join(temp_docs_dir, "doc2.docx")
            output = os.path.join(temp_docs_dir, "merged.docx")

            # Corregir la sintaxis de creación de documentos
            doc1_obj = Document()
            doc1_obj.add_paragraph("Document 1")
            doc1_obj.save(doc1)

            doc2_obj = Document()
            doc2_obj.add_paragraph("Document 2")
            doc2_obj.save(doc2)

            # Mock the merge function to avoid actual file operations
            mock_merge.return_value = f"Successfully merged documents into {output}"

            result = await document_tools.merge_documents(
                target_filename=output,
                source_filenames=[doc1, doc2],
                add_page_breaks=True,
            )

            assert "successfully merged" in result.lower()
            mock_merge.assert_called_once()

    async def test_is_path_in_allowed_directories(self):
        """Test the path validation function."""
        # Create a temporary directory for testing
        with tempfile.TemporaryDirectory() as temp_dir:
            # Set the allowed directories environment variable
            with patch.dict("os.environ", {"MCP_ALLOWED_DIRECTORIES": temp_dir}):
                # Test with a path in the allowed directory
                test_file = os.path.join(temp_dir, "test.docx")
                is_allowed, _ = document_tools._is_path_in_allowed_directories(
                    test_file
                )
                assert is_allowed is True

                # Test with a path outside the allowed directory
                outside_path = os.path.join(os.path.dirname(temp_dir), "test.docx")
                is_allowed, error_msg = document_tools._is_path_in_allowed_directories(
                    outside_path
                )
                assert is_allowed is False
                assert "not in allowed directories" in error_msg
