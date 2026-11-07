"""Header and footer tools shim for backward compatibility.

All functions have been consolidated into word_mcp.tools.format_tools.
"""

from word_mcp.tools.format_tools import add_footer, add_header

__all__ = ["add_footer", "add_header"]
