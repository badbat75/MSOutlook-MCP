"""
Outlook MCP Server - Entry Point
=================================
Thin wrapper that imports and runs the MCP server from the outlook_mcp package.
See outlook_mcp/server.py for the full implementation.

Usage:
    python outlook_mcp_server.py          # stdio transport (for Claude Desktop)
    python outlook_mcp_server.py --http   # HTTP transport (for remote)
"""

from outlook_mcp.server import main

if __name__ == "__main__":
    main()
