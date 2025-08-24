"""
Entry point for running mcp_word as a module.

This module provides the command-line interface for the MCP Word Office server,
supporting both MCP server mode and standalone CLI operations. It handles
environment variable configuration, command-line argument parsing, and provides
clear error messages for configuration issues.

The module supports multiple execution modes:
- MCP Server: `python -m mcp_word server` or `uv run mcp_word_office`
- File listing: `python -m mcp_word list`
- Default behavior: Start MCP server when no command is specified
"""

from mcp_word.server import create_server


def main() -> None:
    mcp = create_server()
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
