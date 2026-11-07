"""Link tools shim for backward compatibility.

All link and bookmark functions have been consolidated into word_mcp.tools.content_tools.
"""

from word_mcp.tools.content_tools import add_bookmark, add_hyperlink

__all__ = ["add_bookmark", "add_hyperlink"]
