"""Tests for exception centralization module."""

import pytest

from mcp_word.exception import (
    ConfigurationError,
    DocumentProcessingError,
    ExceptionTool,
    FileOperationError,
    StyleError,
    ValidationError,
)


class TestExceptionCore:
    """Tests for ExceptionCore exceptions."""

    def test_document_processing_error(self):
        """Test DocumentProcessingError creation and inheritance."""
        error = DocumentProcessingError("Test error message")
        assert isinstance(error, Exception)
        assert str(error) == "Test error message"

    def test_validation_error(self):
        """Test ValidationError inherits from ValueError."""
        error = ValidationError("Invalid input")
        assert isinstance(error, ValueError)
        assert isinstance(error, Exception)
        assert str(error) == "Invalid input"

    def test_file_operation_error(self):
        """Test FileOperationError inherits from IOError."""
        error = FileOperationError("File not found")
        assert isinstance(error, IOError)
        assert isinstance(error, Exception)
        assert str(error) == "File not found"

    def test_style_error(self):
        """Test StyleError creation."""
        error = StyleError("Style not found")
        assert isinstance(error, Exception)
        assert str(error) == "Style not found"

    def test_configuration_error(self):
        """Test ConfigurationError creation."""
        error = ConfigurationError("Invalid config")
        assert isinstance(error, Exception)
        assert str(error) == "Invalid config"


class TestExceptionTool:
    """Tests for ExceptionTool utilities."""

    def test_get_error_info_known_exception(self):
        """Test get_error_info with known exception types."""
        error_info = ExceptionTool.get_error_info(DocumentProcessingError("test"))
        assert error_info["error_type"] == "document_processing"
        assert error_info["recoverable"] is True
        assert error_info["message"] == "test"

    def test_get_error_info_validation_error(self):
        """Test get_error_info with ValidationError."""
        error_info = ExceptionTool.get_error_info(ValidationError("invalid"))
        assert error_info["error_type"] == "validation"
        assert error_info["recoverable"] is False

    def test_get_error_info_file_operation_error(self):
        """Test get_error_info with FileOperationError."""
        error_info = ExceptionTool.get_error_info(FileOperationError("file error"))
        assert error_info["error_type"] == "file_operation"
        assert error_info["recoverable"] is True

    def test_get_error_info_unknown_exception(self):
        """Test get_error_info with unknown exception."""
        error_info = ExceptionTool.get_error_info(RuntimeError("unknown"))
        assert error_info["error_type"] == "unknown"
        assert error_info["recoverable"] is True

    def test_to_operation_error(self):
        """Test conversion of exception to OperationError."""
        op_error = ExceptionTool.to_operation_error(
            DocumentProcessingError("test error"),
            suggestion="Check the document format",
        )
        assert op_error.status == "error"
        assert op_error.error_type == "document_processing"
        assert op_error.message == "test error"
        assert op_error.suggestion == "Check the document format"
        assert op_error.recoverable is True

    def test_handle_error_returns_dict(self):
        """Test handle_error returns standardized dictionary."""
        result = ExceptionTool.handle_error(
            FileOperationError("test"), filename="test.docx", operation="read"
        )
        assert isinstance(result, dict)
        assert result["status"] == "error"
        assert result["error_type"] == "file_operation"
        assert "recoverable" in result

    def test_handle_error_with_filename(self):
        """Test handle_error includes filename in message."""
        result = ExceptionTool.handle_error(
            Exception("test error"), filename="test.docx", operation="read"
        )
        assert "test.docx" in result["message"]
        assert "read" in result["message"]

    def test_handle_error_without_context(self):
        """Test handle_error without optional context."""
        result = ExceptionTool.handle_error(Exception("test error"))
        assert result["message"] == "test error"

    def test_get_suggestion(self):
        """Test _get_suggestion returns appropriate suggestions."""
        suggestion = ExceptionTool._get_suggestion("document_processing", Exception())
        assert suggestion is not None
        assert "document" in suggestion.lower() or "format" in suggestion.lower()

        suggestion = ExceptionTool._get_suggestion("validation", Exception())
        assert suggestion is not None
        assert "input" in suggestion.lower() or "parameter" in suggestion.lower()

        suggestion = ExceptionTool._get_suggestion("file_operation", Exception())
        assert suggestion is not None
        assert "permission" in suggestion.lower() or "file" in suggestion.lower()

    def test_exception_mapping_coverage(self):
        """Test that all core exceptions are mapped."""
        exceptions_to_test = [
            (DocumentProcessingError, "document_processing"),
            (ValidationError, "validation"),
            (FileOperationError, "file_operation"),
            (StyleError, "style"),
            (ConfigurationError, "configuration"),
        ]

        for exc_class, expected_type in exceptions_to_test:
            error_info = ExceptionTool.get_error_info(exc_class("test"))
            assert error_info["error_type"] == expected_type, (
                f"{exc_class.__name__} should map to {expected_type}"
            )


class TestExceptionImportFromPackage:
    """Tests for importing exceptions from package level."""

    def test_import_from_mcp_word(self):
        """Test exceptions can be imported from mcp_word package."""
        from mcp_word import (
            ConfigurationError,
            DocumentProcessingError,
            ExceptionTool,
            FileOperationError,
            StyleError,
            ValidationError,
        )

        assert DocumentProcessingError is not None
        assert ValidationError is not None
        assert FileOperationError is not None
        assert StyleError is not None
        assert ConfigurationError is not None
        assert ExceptionTool is not None


class TestExceptionCentralization:
    """Tests to verify exception centralization works correctly."""

    def test_exceptions_available_from_exception_module(self):
        """Test all exceptions are available from mcp_word.exception."""
        from mcp_word.exception import (
            ConfigurationError,
            DocumentProcessingError,
            FileOperationError,
            StyleError,
            ValidationError,
        )

        error = DocumentProcessingError("test")
        assert str(error) == "test"

        error = ValidationError("test")
        assert isinstance(error, ValueError)

        error = FileOperationError("test")
        assert isinstance(error, IOError)

        error = StyleError("test")
        assert isinstance(error, Exception)

        error = ConfigurationError("test")
        assert isinstance(error, Exception)

    def test_core_exceptions_module_removed(self):
        """Verify core/exceptions.py has been removed."""
        import importlib
        import sys

        assert "mcp_word.core.exceptions" not in sys.modules
        with pytest.raises(ModuleNotFoundError):
            importlib.import_module("mcp_word.core.exceptions")
