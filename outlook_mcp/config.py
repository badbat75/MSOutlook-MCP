"""Server configuration loaded from an external TOML file.

The transport (stdio or http) and the HTTP bind address live in a small TOML
file, never on the command line, so the same entry point behaves the same
whether it is spawned by Claude Desktop, a daemon, or an operator:

    [server]
    transport = "http"        # "stdio" (default) or "http"
    bind_host = "0.0.0.0"     # HTTP only
    bind_port = 8000          # HTTP only

Lookup order for the file:

1. ``--config PATH`` on the command line
2. the ``OUTLOOK_MCP_CONFIG`` environment variable
3. ``outlook_mcp.toml`` next to the project root (the directory holding
   ``outlook_mcp_server.py``), so it does not depend on the process CWD

When no file exists the defaults apply (stdio transport). A file that exists
but cannot be parsed or contains invalid values is an error: silently falling
back to stdio would leave an HTTP deployment unreachable with no hint why.

Credentials never belong in this file. In stdio mode they come from the
``OUTLOOK_*`` environment variables; in HTTP mode from request headers.
"""

import logging
import os
import tomllib
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

logger = logging.getLogger("outlook_mcp")

CONFIG_ENV_VAR = "OUTLOOK_MCP_CONFIG"
CONFIG_FILENAME = "outlook_mcp.toml"
DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent.parent / CONFIG_FILENAME

TRANSPORT_STDIO = "stdio"
TRANSPORT_HTTP = "http"
VALID_TRANSPORTS = (TRANSPORT_STDIO, TRANSPORT_HTTP)

# Path under which the streamable HTTP endpoint is mounted.
HTTP_PATH = "/mcp"


class ConfigError(ValueError):
    """Raised when the configuration file is present but unusable."""


@dataclass(frozen=True)
class ServerConfig:
    """Listening parameters for the MCP server."""

    transport: str = TRANSPORT_STDIO
    bind_host: str = "127.0.0.1"
    bind_port: int = 8000
    source: Optional[Path] = None
    """The file the values were read from, or None when defaults apply."""

    @property
    def is_http(self) -> bool:
        return self.transport == TRANSPORT_HTTP

    @property
    def http_url(self) -> str:
        """URL clients should connect to (HTTP mode only)."""
        host = self.bind_host
        # A wildcard bind is reachable on every interface; show the loopback
        # address rather than a literal 0.0.0.0 that no client can dial.
        if host in ("0.0.0.0", "::"):
            host = "127.0.0.1"
        if ":" in host and not host.startswith("["):
            host = f"[{host}]"
        return f"http://{host}:{self.bind_port}{HTTP_PATH}"


def resolve_config_path(cli_path: Optional[str] = None) -> Optional[Path]:
    """Pick the configuration file to read, or None when none is configured.

    An explicitly requested path (CLI flag or environment variable) must
    exist: a typo there should fail loudly rather than start a stdio server.
    The implicit project-root file is optional.
    """
    explicit = cli_path or os.environ.get(CONFIG_ENV_VAR)
    if explicit:
        path = Path(explicit).expanduser()
        if not path.is_file():
            raise ConfigError(f"Configuration file not found: {path}")
        return path
    if DEFAULT_CONFIG_PATH.is_file():
        return DEFAULT_CONFIG_PATH
    return None


def _parse(data: dict, source: Path) -> ServerConfig:
    server = data.get("server", {})
    if not isinstance(server, dict):
        raise ConfigError(f"{source}: [server] must be a table")

    known = {"transport", "bind_host", "bind_port"}
    for key in server:
        if key not in known:
            logger.warning("%s: ignoring unknown key [server].%s", source, key)

    transport = server.get("transport", TRANSPORT_STDIO)
    if not isinstance(transport, str) or transport.lower() not in VALID_TRANSPORTS:
        raise ConfigError(
            f"{source}: [server].transport must be one of {', '.join(VALID_TRANSPORTS)} "
            f"(got {transport!r})"
        )
    transport = transport.lower()

    bind_host = server.get("bind_host", "127.0.0.1")
    if not isinstance(bind_host, str) or not bind_host.strip():
        raise ConfigError(f"{source}: [server].bind_host must be a non-empty string")

    bind_port = server.get("bind_port", 8000)
    # bool is a subclass of int; reject it explicitly so `bind_port = true`
    # does not silently become port 1.
    if isinstance(bind_port, bool) or not isinstance(bind_port, int) or not 1 <= bind_port <= 65535:
        raise ConfigError(f"{source}: [server].bind_port must be an integer between 1 and 65535")

    return ServerConfig(
        transport=transport,
        bind_host=bind_host.strip(),
        bind_port=bind_port,
        source=source,
    )


def load_config(cli_path: Optional[str] = None) -> ServerConfig:
    """Load the server configuration, falling back to stdio defaults."""
    path = resolve_config_path(cli_path)
    if path is None:
        logger.info(
            "No %s found (looked for %s); using stdio transport", CONFIG_FILENAME, DEFAULT_CONFIG_PATH
        )
        return ServerConfig()
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
    except OSError as e:
        raise ConfigError(f"Cannot read configuration file {path}: {e}") from e
    except tomllib.TOMLDecodeError as e:
        raise ConfigError(f"Invalid TOML in {path}: {e}") from e
    return _parse(data, path)
