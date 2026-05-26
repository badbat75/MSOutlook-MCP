"""
Outlook MCP Server - Entry Point
=================================
Thin wrapper that imports and runs the MCP server from the outlook_mcp package.
See outlook_mcp/server.py for the full implementation.

Usage:
    python outlook_mcp_server.py          # stdio transport (for Claude Desktop)
    python outlook_mcp_server.py --http   # HTTP transport (for remote)
"""

import os
from pathlib import Path

# ── Load credentials from .env (single source of truth) ──────────────
_env_path = Path(__file__).resolve().parent / ".env"
if _env_path.exists():
    for line in _env_path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip()
        # Only set if not already set by environment
        if key not in os.environ:
            os.environ[key] = val

from outlook_mcp.server import main

if __name__ == "__main__":
    main()
