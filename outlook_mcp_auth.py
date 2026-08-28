#!/usr/bin/env python3
"""
Outlook MCP - OAuth2 Authentication Setup
==========================================
Runs the authorization flow from the outlook_mcp package. Exactly equivalent to
the ``outlook-mcp-auth`` console script installed by ``pip install .``: both
call ``outlook_mcp.authorize.main()``.

Usage:
    python outlook_mcp_auth.py                     # opens a browser
    python outlook_mcp_auth.py --no-browser        # headless / SSH
    python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'
    python outlook_mcp_auth.py --user a@b.com      # one user of an HTTP deployment

The app registration comes from the same outlook_mcp.toml the server reads.
See outlook_mcp/authorize.py for the full implementation.
"""

from outlook_mcp.authorize import main

if __name__ == "__main__":
    main()
