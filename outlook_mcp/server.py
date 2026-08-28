"""
Outlook MCP Server - entry point.

Reads the .env and the TOML configuration, then starts the transport the
configuration names. The 19 tools themselves live in outlook_mcp.tools and
register themselves on the server object in outlook_mcp.app when imported.
"""

import argparse
import sys
from typing import List, Optional

from .app import mcp, set_config
from .config import ConfigError, load_config
from .env import EnvFileError, load_project_env

# Imported for the side effect: this is what puts the tools on `mcp`.
from . import tools  # noqa: F401


def _parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="outlook_mcp_server",
        description=(
            "Outlook MCP server. The transport (stdio or http) and the HTTP bind "
            "address are read from outlook_mcp.toml, not from the command line."
        ),
    )
    parser.add_argument(
        "--config",
        metavar="PATH",
        help=(
            "Path to the TOML configuration file. Defaults to $OUTLOOK_MCP_CONFIG, "
            "then outlook_mcp.toml in the project root."
        ),
    )
    parser.add_argument(
        "--env-file",
        metavar="PATH",
        help=(
            "Path to the .env holding the OUTLOOK_* credentials. Defaults to "
            "$OUTLOOK_ENV_FILE, then .env in the project root. Variables already "
            "set in the environment always win."
        ),
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None):
    """Run the MCP server with the transport chosen by the configuration file."""
    args = _parse_args(argv)
    try:
        # Before anything reads OUTLOOK_*: this is what makes the console
        # script and the outlook_mcp_server.py wrapper behave identically.
        load_project_env(args.env_file)
        config = load_config(args.config)
    except (ConfigError, EnvFileError) as e:
        print(f"Configuration error: {e}", file=sys.stderr)
        sys.exit(2)

    set_config(config)

    if config.is_http:
        # stderr on purpose: stdout is never used by the HTTP transport, but
        # keeping all diagnostics on one stream makes daemon logs predictable.
        print(
            f"Starting Outlook MCP server on {config.http_url} "
            f"(bind {config.bind_host}:{config.bind_port}, config: {config.source})",
            file=sys.stderr,
        )
        # No transport_security override: for loopback binds the SDK enables
        # DNS-rebinding protection for localhost hosts on its own, and for any
        # other bind address it leaves Host checking off, which is what a
        # server reached through a hostname or reverse proxy needs.
        mcp.run(
            transport="streamable-http",
            host=config.bind_host,
            port=config.bind_port,
        )
    else:
        mcp.run()  # stdio transport (Claude Desktop, Claude Code, daemons)


if __name__ == "__main__":
    main()
