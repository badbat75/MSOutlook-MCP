#!/usr/bin/env python3
"""
Outlook MCP Server - Entry Point
=================================
Runs the MCP server from the outlook_mcp package. Exactly equivalent to the
``outlook-mcp`` console script installed by ``pip install .``: both call
``outlook_mcp.server.main()``, which loads the project .env, reads the
configuration file and starts the transport it names.

Usage:
    python outlook_mcp_server.py                      # transport from outlook_mcp.toml
    python outlook_mcp_server.py --config /etc/x.toml # explicit configuration file
    python outlook_mcp_server.py --env-file /etc/x.env

The transport (stdio or http) and the HTTP bind host/port are read from the
TOML configuration file (see outlook_mcp.toml.example), never from flags.
Credentials come from OUTLOOK_* environment variables in stdio mode and from
the X-Outlook-* request headers in HTTP mode.
"""

from outlook_mcp.server import main

if __name__ == "__main__":
    main()
