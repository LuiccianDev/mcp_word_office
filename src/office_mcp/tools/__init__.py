"""
MCP tool implementations for Office MCP Server.

This package contains the MCP tool implementations that expose functionality
to clients through the Model Context Protocol.
"""

from typing import Any


def __getattr__(name: str) -> Any:
    if name == "register_all_tools":
        from office_mcp.tools.register_tools import register_all_tools
        return register_all_tools
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = [
    "register_all_tools",
]
