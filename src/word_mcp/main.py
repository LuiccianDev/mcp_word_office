"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
"""

# modulos de terceros
from mcp.server.fastmcp import FastMCP

from word_mcp.prompts.register_prompts import register_prompts

# modulos propios
from word_mcp.tools.register_tools import register_all_tools


def create_server() -> FastMCP:
    """Run the Word Document MCP Server."""
    # Initialize FastMCP server
    mcp = FastMCP(
        name="word-document-server",
        instructions="Servidor MCP para informes Word en minería.",
        dependencies=["python-docx"],
        on_duplicate_tools="error",
    )
    # Register all tools
    register_all_tools(mcp)

    # Register all prompts
    register_prompts(mcp)

    return mcp
