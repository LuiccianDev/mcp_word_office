"""Header and footer tools shim for backward compatibility.

All functions have been consolidated into office_mcp.tools.format_tools.
"""

from office_mcp.tools.format_tools import add_footer, add_header

__all__ = ["add_footer", "add_header"]
