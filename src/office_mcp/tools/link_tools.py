"""Link tools shim for backward compatibility.

All link and bookmark functions have been consolidated into office_mcp.tools.content_tools.
"""

from office_mcp.tools.content_tools import add_bookmark, add_hyperlink

__all__ = ["add_bookmark", "add_hyperlink"]
