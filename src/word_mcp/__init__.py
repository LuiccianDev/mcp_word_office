"""
Word MCP Server

A Model Context Protocol (MCP) server providing comprehensive Word (.docx) manipulation
capabilities for AI assistants and applications. This package offers programmatic access
to Word operations including text extraction, formatting, styles, and document structure
integration through a clean, well-documented API.
"""

from importlib.metadata import PackageNotFoundError, version


try:
    __version__ = version("luiccian-word-mcp")
except PackageNotFoundError:
    __version__ = "0.0.0+unknown"

__author__ = "LuiccianDev"
__description__ = "Word MCP Server - High-performance Word document manipulation via Model Context Protocol"
__title__ = "word_mcp"

from word_mcp.exception import (
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
