"""
Office MCP Server

A Model Context Protocol (MCP) server providing comprehensive Word (.docx) manipulation
capabilities for AI assistants and applications. This package offers programmatic access
to Word operations including text extraction, formatting, styles, and document structure
integration through a clean, well-documented API.
"""

__version__ = "1.1.2"
__author__ = "LuiccianDev"
__description__ = (
    "Office MCP Server - Word document manipulation through Model Context Protocol"
)
__title__ = "office_mcp"

from office_mcp.exception import (
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
    "ConfigurationError",
    "DocumentProcessingError",
    "ExceptionTool",
    "FileOperationError",
    "StyleError",
    "ValidationError",
]
