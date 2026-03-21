"""
MCP tool implementations for the Word Document Server.

This package contains the MCP tool implementations that expose functionality
to clients through the Model Context Protocol.
"""

from mcp_word.tools.register_tools import register_all_tools


__all__ = [
    "register_all_tools",
]
