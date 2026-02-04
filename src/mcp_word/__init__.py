"""
MCP Word Office Server

A Model Context Protocol (MCP) server providing comprehensive Word (.docx) manipulation
capabilities for AI assistants and applications. This package offers programmatic access
to Word operations including text extraction, formatting, styles, and document structure
integration through a clean, well-documented API.
"""

__version__ = "1.0.1"
__author__ = "LuiccianDev"
__description__ = (
    "MCP Word Office Server - Word document manipulation through Model Context Protocol"
)
__title__ = "mcp_word"

from .exception import (
    ConfigurationError,
    DocumentProcessingError,
    ExceptionTool,
    FileOperationError,
    StyleError,
    ValidationError,
)


__all__ = [
    "__version__",
    "__author__",
    "__description__",
    "__title__",
    "DocumentProcessingError",
    "ValidationError",
    "FileOperationError",
    "StyleError",
    "ConfigurationError",
    "ExceptionTool",
]
