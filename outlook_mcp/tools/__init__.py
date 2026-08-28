"""Tool definitions.

Importing this package is what registers the tools: each module runs its
``@mcp.tool()`` decorators against the shared server in outlook_mcp.app.
server.py imports it for that side effect alone.
"""

from . import calendar, mail, profile  # noqa: F401
