"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
"""

from mcp.server.fastmcp import FastMCP
from mcp_word_server.tools import register_all_tools

def create_server() -> FastMCP:
    """Run the Word Document MCP Server."""
    # Initialize FastMCP server
    mcp = FastMCP("word-document-server")
    # Register all tools
    register_all_tools(mcp)

    return mcp

