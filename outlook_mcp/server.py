"""
Outlook MCP Server - entry point.

Reads the TOML configuration, then starts the transport it names. The 19 tools
themselves live in outlook_mcp.tools and register themselves on the server
object in outlook_mcp.app when imported.
"""

import argparse
import sys
from typing import List, Optional

from mcp.server.transport_security import TransportSecuritySettings

from .app import mcp, set_config
from .config import ConfigError, ServerConfig, load_config

# Imported for the side effect: this is what puts the tools on `mcp`, and the
# enrollment and download routes on its HTTP app. The routes answer 404 unless
# the loaded configuration enables them, so importing them unconditionally is
# safe and keeps registration independent of when the configuration is installed.
from . import downloads, enroll, tools  # noqa: F401


def _parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="outlook_mcp_server",
        description=(
            "Outlook MCP server. Everything it needs, transport included, is read "
            "from outlook_mcp.toml, not from the command line or the environment."
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
    return parser.parse_args(argv)


def _transport_security(config: ServerConfig) -> Optional[TransportSecuritySettings]:
    """What Host and Origin headers the HTTP transport should accept.

    None when [server].allowed_hosts is empty, which is identical to passing
    nothing: the SDK then applies its own rule, and for a loopback bind that
    means DNS-rebinding protection with localhost as the only acceptable Host
    (mcp/server/lowlevel/server.py). That rule assumes a proxied server binds a
    routable address, and this one does not: it binds loopback *and* sits behind
    a proxy, so a request arrives carrying the proxy's own hostname and is
    answered 421 Invalid Host header.

    Naming the hostname here keeps the protection on rather than turning it off,
    which is why this widens the list instead of disabling the check.
    """
    if not config.allowed_hosts:
        return None

    # Both forms of each name: nginx's default `proxy_set_header Host $host`
    # forwards the bare hostname, but a proxy configured with $http_host passes
    # the port along, and the SDK matches the Host header literally.
    hosts = [form for host in config.allowed_hosts for form in (host, f"{host}:*")]
    return TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=[*hosts, "127.0.0.1:*", "localhost:*", "[::1]:*"],
        allowed_origins=[
            *(f"https://{host}" for host in config.allowed_hosts),
            "http://127.0.0.1:*", "http://localhost:*", "http://[::1]:*",
        ],
    )


def main(argv: Optional[List[str]] = None):
    """Run the MCP server with the transport chosen by the configuration file."""
    args = _parse_args(argv)
    try:
        config = load_config(args.config)
    except ConfigError as e:
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
        mcp.run(
            transport="streamable-http",
            host=config.bind_host,
            port=config.bind_port,
            transport_security=_transport_security(config),
        )
    else:
        mcp.run()  # stdio transport (Claude Desktop, Claude Code, daemons)


if __name__ == "__main__":
    main()
