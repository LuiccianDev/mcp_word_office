#!/usr/bin/env python3
"""
Run script for the Word Document Server.

This script provides a simple way to start the Word Document Server.
"""

from mcp_word_server.main import mcp

if __name__ == "__main__":
    mcp.run(transport='stdio')
