#!/usr/bin/env python3
"""
Outlook MCP Server - Entry Point
=================================
Runs the MCP server from the outlook_mcp package. Exactly equivalent to the
``outlook-mcp`` console script installed by ``pip install .``: both call
``outlook_mcp.server.main()``, which reads the configuration file and starts
the transport it names.

Usage:
    python outlook_mcp_server.py                      # transport from outlook_mcp.toml
    python outlook_mcp_server.py --config /etc/x.toml # explicit configuration file

Everything the server needs is in that TOML file (see outlook_mcp.toml.example):
the transport, the HTTP bind host and port, and the Azure AD app registration.
No credential is ever read from the environment or taken from a request.
"""

from outlook_mcp.server import main

if __name__ == "__main__":
    main()
