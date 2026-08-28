#!/usr/bin/env python3
"""
Outlook MCP - OAuth2 Authentication Setup
==========================================
Runs the authorization flow from the outlook_mcp package. Exactly equivalent to
the ``outlook-mcp-auth`` console script installed by ``pip install .``: both
call ``outlook_mcp.authorize.main()``.

Usage:
    python outlook_mcp_auth.py                # opens a browser
    python outlook_mcp_auth.py --no-browser   # headless / SSH
    python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'

Credentials come from the project .env, or from the environment if already set.
See outlook_mcp/authorize.py for the full implementation.
"""

from outlook_mcp.authorize import main

if __name__ == "__main__":
    main()
