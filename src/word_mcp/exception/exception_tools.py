"""Exception handling utilities for MCP Word Server.
This module provides tools and utilities for consistent exception handling
across all MCP tools, including decorators, error mapping, and response builders.
"""

from collections.abc import Callable
from functools import wraps
from typing import Any

from ..models.response_models import OperationError
from .exception_core import (
    ConfigurationError,
    DocumentProcessingError,
    FileOperationError,
    StyleError,
    ValidationError,
)


class ExceptionTool:
    """Utility class for centralized exception handling."""

    EXCEPTION_MAPPING: dict[type[Exception], dict[str, Any]] = {
        DocumentProcessingError: {
            "error_type": "document_processing",
            "recoverable": True,
        },
        ValidationError: {
            "error_type": "validation",
            "recoverable": False,
        },
        FileOperationError: {
            "error_type": "file_operation",
            "recoverable": True,
        },
        StyleError: {
            "error_type": "style",
            "recoverable": True,
        },
        ConfigurationError: {
            "error_type": "configuration",
            "recoverable": False,
        },
        ValueError: {
            "error_type": "validation",
            "recoverable": False,
        },
        IOError: {
            "error_type": "file_operation",
            "recoverable": True,
        },
        OSError: {
            "error_type": "file_operation",
            "recoverable": True,
        },
        FileNotFoundError: {
            "error_type": "file_not_found",
            "recoverable": False,
        },
        PermissionError: {
            "error_type": "permission_denied",
            "recoverable": False,
        },
    }

    @staticmethod
    def get_error_info(exception: Exception) -> dict[str, Any]:
        """Get error information for an exception.

        Args:
            exception: The exception to analyze.

        Returns:
            Dictionary with error_type, recoverable, and message.
        """
        exc_type = type(exception)

        for exc_class, info in ExceptionTool.EXCEPTION_MAPPING.items():
            if issubclass(exc_type, exc_class):
                return {
                    "error_type": info["error_type"],
                    "recoverable": info["recoverable"],
                    "message": str(exception),
                }

        return {
            "error_type": "unknown",
            "recoverable": True,
            "message": str(exception),
        }

    @staticmethod
    def to_operation_error(
        exception: Exception,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> OperationError:
        """Convert an exception to an OperationError response.

        Args:
            exception: The exception to convert.
            suggestion: Optional suggestion for the user.
            details: Optional additional details.

        Returns:
            OperationError model instance.
        """
        error_info = ExceptionTool.get_error_info(exception)

        return OperationError(
            status="error",
            error_type=error_info["error_type"],
            message=error_info["message"],
            suggestion=suggestion,
            recoverable=error_info["recoverable"],
            details=details,
        )

    @staticmethod
    def handle_error(
        exception: Exception,
        filename: str | None = None,
        operation: str | None = None,
    ) -> dict[str, Any]:
        """Handle an exception and return a standardized error response.

        Args:
            exception: The exception that occurred.
            filename: Optional filename involved in the error.
            operation: Optional description of the operation.

        Returns:
            Dictionary representation of the error response.
        """
        error_info = ExceptionTool.get_error_info(exception)

        context = []
        if filename:
            context.append(f"File: {filename}")
        if operation:
            context.append(f"Operation: {operation}")

        message = error_info["message"]
        if context:
            message = f"{message} ({', '.join(context)})"

        suggestion = ExceptionTool._get_suggestion(error_info["error_type"], exception)

        return {
            "status": "error",
            "error_type": error_info["error_type"],
            "message": message,
            "suggestion": suggestion,
            "recoverable": error_info["recoverable"],
        }

    @staticmethod
    def _get_suggestion(error_type: str, exception: Exception) -> str | None:
        """Get a helpful suggestion based on error type.

        Args:
            error_type: The type of error.
            exception: The exception instance.

        Returns:
            Helpful suggestion string or None.
        """
        suggestions = {
            "document_processing": "Check if the document is corrupted or in a valid format.",
            "validation": "Review the input parameters and ensure they meet the requirements.",
            "file_operation": "Verify file permissions and that the file is not open in another program.",
            "file_not_found": "Check that the file path is correct and the file exists.",
            "permission_denied": "Ensure you have the necessary permissions to access the file.",
            "style": "Verify the style name exists in the document template.",
            "configuration": "Check the application configuration settings.",
        }
        return suggestions.get(error_type)

    @staticmethod
    def wrap_tool_call(
        filename_param: str = "filename",
        suggestion: str | None = None,
    ) -> Callable:
        """Decorator to wrap tool functions with consistent error handling.

        Args:
            filename_param: Name of the filename parameter.
            suggestion: Default suggestion for errors.

        Returns:
            Decorated function.
        """

        def decorator(func: Callable) -> Callable:
            @wraps(func)
            def wrapper(*args: Any, **kwargs: Any) -> Any:
                filename = kwargs.get(filename_param, "unknown")
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    return ExceptionTool.handle_error(e, filename=filename)

            return wrapper

        return decorator
