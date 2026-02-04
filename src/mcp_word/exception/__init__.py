"""Exception handling module for MCP Word Server.

This module provides centralized exception management for the application,
including custom exceptions and utilities for error handling.
"""

from .ExceptionCore import (
    ConfigurationError,
    DocumentProcessingError,
    FileOperationError,
    StyleError,
    ValidationError,
)
from .ExceptionTool import ExceptionTool


__all__ = [
    "DocumentProcessingError",
    "ValidationError",
    "FileOperationError",
    "StyleError",
    "ConfigurationError",
    "ExceptionTool",
]
