#!/usr/bin/env python3
"""
Run script for the Word Document Server.

This script provides a simple way to start the Word Document Server.
"""

from word_mcp.main import create_server

if __name__ == "__main__":
    mcp = create_server()
    mcp.run(transport="stdio")
