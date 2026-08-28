"""The server configuration, and the only place it comes from.

One TOML file holds everything a deployment needs: the transport, the HTTP bind
address, the Azure AD app registration the server runs as, and where downloaded
attachments go. Nothing is taken from the command line beyond the path of this
file, and nothing is taken from the environment at all, so the same entry point
behaves the same whether it is spawned by Claude Desktop, a daemon or an
operator, and a deployment can be read off one artefact:

    [server]
    transport = "http"        # "stdio" (default) or "http"
    bind_host = "127.0.0.1"   # HTTP only, and loopback only
    bind_port = 8000          # HTTP only
    allowed_hosts = ["proxy.example.com"]   # HTTP only; the Host a proxy forwards

    [credentials]
    client_id = "..."
    client_secret = "..."
    tenant_id = "common"

    [auth]
    user_header = "X-Auth-Email"                      # HTTP only
    public_url = "https://outlook-mcp.example.com"    # HTTP only
    cache_dir = "/opt/outlook-mcp/data/caches"        # either transport

    [attachments]
    download_path = "~/Downloads/outlook_attachments"
    retention_minutes = 60    # HTTP only; 0 keeps files until a tool deletes them

Lookup order for the file:

1. ``--config PATH`` on the command line
2. the ``OUTLOOK_MCP_CONFIG`` environment variable
3. ``outlook_mcp.toml`` next to the project root (the directory holding
   ``outlook_mcp_server.py``), so it does not depend on the process CWD

That variable locates the file; it never carries configuration itself. When no
file exists the defaults apply, which is a stdio server with no credentials: it
starts, and says on the first tool call what is missing. A file that exists but
cannot be parsed, or that contains invalid values, is a hard error instead:
silently falling back to stdio would leave an HTTP deployment unreachable with
no hint why.

**No credential ever travels in a request.** Over HTTP a caller says only which
mailbox it is addressing, in the header a reverse proxy fills in for it, and the
server answers with its own credentials and that user's token cache.
"""

import ipaddress
import logging
import os
import tomllib
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Optional

logger = logging.getLogger("outlook_mcp")

PROJECT_ROOT = Path(__file__).resolve().parent.parent
"""Directory holding outlook_mcp_server.py, resolved from this file's location.

Every path this project resolves hangs off here rather than off the CWD: a
stdio server inherits the working directory of whatever host spawned it, which
may not even be traversable by the user the server runs as.
"""

CONFIG_ENV_VAR = "OUTLOOK_MCP_CONFIG"
CONFIG_FILENAME = "outlook_mcp.toml"
DEFAULT_CONFIG_PATH = PROJECT_ROOT / CONFIG_FILENAME

TRANSPORT_STDIO = "stdio"
TRANSPORT_HTTP = "http"
VALID_TRANSPORTS = (TRANSPORT_STDIO, TRANSPORT_HTTP)

DEFAULT_TENANT_ID = "common"
DEFAULT_USER_HEADER = "X-Auth-Email"
DEFAULT_DOWNLOAD_PATH = Path.home() / "Downloads" / "outlook_attachments"

# How long an attachment nobody downloaded stays on the server, in minutes.
DEFAULT_RETENTION_MINUTES = 60

# Path under which the streamable HTTP endpoint is mounted.
HTTP_PATH = "/mcp"

# Paths of the enrollment routes, relative to public_url.
ENROLL_LOGIN_PATH = "/oauth/login"
ENROLL_CALLBACK_PATH = "/oauth/callback"

# Where a remote caller fetches an attachment the server downloaded for it. The
# token is minted by the server and names one file for one user; see downloads.py.
DOWNLOAD_PATH_PREFIX = "/attachments"
DOWNLOAD_ROUTE = f"{DOWNLOAD_PATH_PREFIX}/{{token}}"


class ConfigError(ValueError):
    """Raised when the configuration file is present but unusable."""


@dataclass(frozen=True)
class ServerConfig:
    """Everything a running server was told about itself."""

    transport: str = TRANSPORT_STDIO
    bind_host: str = "127.0.0.1"
    bind_port: int = 8000
    allowed_hosts: tuple[str, ...] = ()
    """Host header values to accept besides loopback, for a reverse proxy that
    forwards its own hostname. Empty leaves the SDK default alone: a loopback
    bind auto-enables DNS-rebinding protection and then accepts localhost only.
    A tuple rather than a list because this dataclass is frozen."""
    client_id: Optional[str] = None
    client_secret: Optional[str] = None
    tenant_id: str = DEFAULT_TENANT_ID
    user_header: str = DEFAULT_USER_HEADER
    public_url: Optional[str] = None
    """Absolute URL clients reach this server on, needed by the enrollment routes."""
    cache_dir: Optional[str] = None
    """Where the MSAL token caches go. None leaves them under the home directory."""
    download_path: Optional[str] = None
    retention_minutes: int = DEFAULT_RETENTION_MINUTES
    """Minutes an attachment nobody downloaded stays on the server. 0 keeps it
    until a tool deletes it."""
    source: Optional[Path] = None
    """The file the values were read from, or None when defaults apply."""

    @property
    def is_http(self) -> bool:
        return self.transport == TRANSPORT_HTTP

    @property
    def has_credentials(self) -> bool:
        return bool(self.client_id and self.client_secret)

    @property
    def attachment_dir(self) -> Path:
        """Where attachments are written. Server side on purpose: in HTTP mode a
        remote caller must never get to choose where the server writes files."""
        if not self.download_path:
            return DEFAULT_DOWNLOAD_PATH
        return Path(self.download_path).expanduser()

    @property
    def retention_seconds(self) -> Optional[int]:
        """How long an unfetched download lives, or None when it never expires.

        None over stdio, and that is not a default worth making configurable:
        there the file the tool wrote is the answer itself, in the caller's own
        download directory, so an expiry would delete the user's file. Over HTTP
        the same file is a staging buffer nobody can reach except through a
        one-time link, and one nobody followed is only litter.
        """
        if not self.is_http or self.retention_minutes <= 0:
            return None
        return self.retention_minutes * 60

    @property
    def cache_directory(self) -> Optional[Path]:
        """Where the token caches go, or None for the home directory default.

        None rather than a resolved default on purpose: the default lives in
        auth.py next to the paths it builds, so this can be handed straight to
        user_cache_path() without every caller repeating the fallback.
        """
        if not self.cache_dir:
            return None
        return Path(self.cache_dir).expanduser()

    @property
    def enrollment_enabled(self) -> bool:
        """Whether the browser enrollment routes can be served.

        They need an absolute URL to hand Entra as the redirect URI, and they
        only make sense over HTTP, where a proxy has established who the visitor
        is: without both, users are enrolled with `outlook-mcp-auth --user`.
        """
        return self.is_http and self.public_url is not None

    @property
    def enroll_callback_url(self) -> Optional[str]:
        """The redirect URI to register on the Entra app registration."""
        if self.public_url is None:
            return None
        return self.public_url.rstrip("/") + ENROLL_CALLBACK_PATH

    @property
    def downloads_enabled(self) -> bool:
        """Whether a downloaded attachment can be handed to the caller.

        The same two conditions as enrollment, for the same reason. Over stdio
        there is nothing to hand over: the caller shares the filesystem and the
        path in the answer is already the file. Over HTTP the answer has to
        carry an absolute URL, and only the operator knows what clients reach
        this server on, so without public_url the file stays on the server and
        the tool says so.
        """
        return self.is_http and self.public_url is not None

    def download_url(self, token: str) -> Optional[str]:
        """The URL one download token is redeemed at."""
        if self.public_url is None:
            return None
        return f"{self.public_url.rstrip('/')}{DOWNLOAD_PATH_PREFIX}/{token}"

    @property
    def http_url(self) -> str:
        """URL clients should connect to (HTTP mode only)."""
        host = self.bind_host
        if ":" in host and not host.startswith("["):
            host = f"[{host}]"
        return f"http://{host}:{self.bind_port}{HTTP_PATH}"


def is_loopback(host: str) -> bool:
    """Whether a bind address accepts connections from this machine only."""
    cleaned = host.strip().strip("[]")
    if cleaned in ("localhost", "localhost."):
        return True
    try:
        return ipaddress.ip_address(cleaned).is_loopback
    except ValueError:
        # A hostname we cannot resolve here. Treat it as reachable from
        # elsewhere: assuming otherwise is the mistake that exposes mailboxes.
        return False


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


def _table(data: dict, name: str, source: Path) -> dict:
    """One top-level table, checked to be a table."""
    value = data.get(name, {})
    if not isinstance(value, dict):
        raise ConfigError(f"{source}: [{name}] must be a table")
    return value


def _warn_unknown(table: dict, known: set, name: str, source: Path) -> None:
    for key in table:
        if key not in known:
            logger.warning("%s: ignoring unknown key [%s].%s", source, name, key)


def _string(table: dict, key: str, table_name: str, source: Path) -> Optional[str]:
    """An optional non-empty string value."""
    value = table.get(key)
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"{source}: [{table_name}].{key} must be a non-empty string")
    return value.strip()


def _parse_server(data: dict, source: Path) -> dict:
    server = _table(data, "server", source)
    _warn_unknown(
        server, {"transport", "bind_host", "bind_port", "allowed_hosts"}, "server", source
    )

    transport = server.get("transport", TRANSPORT_STDIO)
    if not isinstance(transport, str) or transport.lower() not in VALID_TRANSPORTS:
        raise ConfigError(
            f"{source}: [server].transport must be one of {', '.join(VALID_TRANSPORTS)} "
            f"(got {transport!r})"
        )

    bind_host = server.get("bind_host", "127.0.0.1")
    if not isinstance(bind_host, str) or not bind_host.strip():
        raise ConfigError(f"{source}: [server].bind_host must be a non-empty string")

    bind_port = server.get("bind_port", 8000)
    # bool is a subclass of int; reject it explicitly so `bind_port = true`
    # does not silently become port 1.
    if isinstance(bind_port, bool) or not isinstance(bind_port, int) or not 1 <= bind_port <= 65535:
        raise ConfigError(f"{source}: [server].bind_port must be an integer between 1 and 65535")

    allowed_hosts = server.get("allowed_hosts", [])
    if not isinstance(allowed_hosts, list) or not all(
        isinstance(host, str) and host.strip() for host in allowed_hosts
    ):
        raise ConfigError(
            f"{source}: [server].allowed_hosts must be a list of non-empty strings"
        )
    for host in allowed_hosts:
        # A Host header is a name and an optional port, never a URL. Getting
        # this wrong costs a 421 that says nothing about the reason, so refuse
        # it here where the message can.
        if "://" in host or "/" in host:
            raise ConfigError(
                f"{source}: [server].allowed_hosts takes Host header values, not "
                f"URLs (got {host!r}). Write the hostname on its own, as it "
                f"appears in the Host header: {host.split('://')[-1].split('/')[0]!r}"
            )

    return {
        "transport": transport.lower(),
        "bind_host": bind_host.strip(),
        "bind_port": bind_port,
        "allowed_hosts": tuple(host.strip() for host in allowed_hosts),
    }


def _parse_credentials(data: dict, source: Path) -> dict:
    creds = _table(data, "credentials", source)
    _warn_unknown(creds, {"client_id", "client_secret", "tenant_id"}, "credentials", source)

    client_id = _string(creds, "client_id", "credentials", source)
    client_secret = _string(creds, "client_secret", "credentials", source)
    if bool(client_id) != bool(client_secret):
        # Half a credential set is always a mistake, and one that would
        # otherwise surface as an opaque failure on the first Graph call.
        raise ConfigError(
            f"{source}: [credentials] needs both client_id and client_secret, "
            f"or neither"
        )

    return {
        "client_id": client_id,
        "client_secret": client_secret,
        "tenant_id": _string(creds, "tenant_id", "credentials", source) or DEFAULT_TENANT_ID,
    }


def _is_rooted(value: str) -> bool:
    """Whether a path is anchored somewhere, rather than at the working directory.

    Judged for both conventions, not just the one this process happens to run
    on: a Linux deployment's configuration is routinely written and validated
    from Windows, where Path("/opt/x").is_absolute() is False for want of a
    drive letter. What the check is really for is refusing a path that would
    follow the CWD, and "/opt/x" never does that on the host that will use it.
    """
    expanded = Path(value).expanduser()
    return expanded.is_absolute() or PurePosixPath(expanded).is_absolute()


def _parse_auth(data: dict, source: Path) -> dict:
    auth = _table(data, "auth", source)
    _warn_unknown(auth, {"user_header", "public_url", "cache_dir"}, "auth", source)

    public_url = _string(auth, "public_url", "auth", source)
    if public_url is not None and not public_url.startswith(("http://", "https://")):
        raise ConfigError(
            f"{source}: [auth].public_url must be an absolute http(s) URL "
            f"(got {public_url!r})"
        )

    cache_dir = _string(auth, "cache_dir", "auth", source)
    if cache_dir is not None and not _is_rooted(cache_dir):
        raise ConfigError(
            f"{source}: [auth].cache_dir must be an absolute path (got "
            f"{cache_dir!r}). A relative one would follow the working directory, "
            f"and a stdio server inherits whatever directory spawned it."
        )

    return {
        "user_header": _string(auth, "user_header", "auth", source) or DEFAULT_USER_HEADER,
        "public_url": public_url,
        "cache_dir": cache_dir,
    }


def _parse_attachments(data: dict, source: Path) -> dict:
    attachments = _table(data, "attachments", source)
    _warn_unknown(
        attachments, {"download_path", "retention_minutes"}, "attachments", source
    )

    retention = attachments.get("retention_minutes", DEFAULT_RETENTION_MINUTES)
    # bool is a subclass of int, and `retention_minutes = false` reading as
    # "0, keep them forever" is the opposite of what anyone writing it meant.
    if isinstance(retention, bool) or not isinstance(retention, int) or retention < 0:
        raise ConfigError(
            f"{source}: [attachments].retention_minutes must be a non-negative "
            f"integer (0 keeps downloads until a tool deletes them)"
        )

    return {
        "download_path": _string(attachments, "download_path", "attachments", source),
        "retention_minutes": retention,
    }


def _validate_deployment(config: ServerConfig, source: Path) -> None:
    """Refuse an HTTP deployment where the user header could be forged.

    Over HTTP the request carries no credential: the user header is trustworthy
    only because the reverse proxy replaces whatever the client sent. If the
    server is reachable without going through that proxy, anyone who can open a
    socket to it can name any mailbox and be served that user's tokens. Making
    the proxy the only route in is therefore not a hardening step but the basis
    of the whole mode, and the one way to establish it from here is to refuse to
    listen anywhere but loopback.
    """
    if not config.is_http or is_loopback(config.bind_host):
        return
    raise ConfigError(
        f"{source}: [server].bind_host must be a loopback address, not "
        f"{config.bind_host!r}. Over HTTP a request proves nothing about itself: the "
        f"{config.user_header} header is only worth believing because the reverse proxy "
        f"replaces whatever the client sent, so anyone able to reach this port directly "
        f"could name any mailbox and be served that user's tokens. Bind to 127.0.0.1 and "
        f"let the proxy be the only way in; if the proxy runs on another host, reach the "
        f"server over a tunnel rather than opening the port."
    )


def _parse(data: dict, source: Path) -> ServerConfig:
    return ServerConfig(
        source=source,
        **_parse_server(data, source),
        **_parse_credentials(data, source),
        **_parse_auth(data, source),
        **_parse_attachments(data, source),
    )


def load_config(cli_path: Optional[str] = None) -> ServerConfig:
    """Load the server configuration, falling back to stdio defaults."""
    path = resolve_config_path(cli_path)
    if path is None:
        logger.info(
            "No %s found (looked for %s); using stdio transport with no credentials",
            CONFIG_FILENAME, DEFAULT_CONFIG_PATH,
        )
        return ServerConfig()
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
    except OSError as e:
        raise ConfigError(f"Cannot read configuration file {path}: {e}") from e
    except tomllib.TOMLDecodeError as e:
        raise ConfigError(f"Invalid TOML in {path}: {e}") from e
    config = _parse(data, path)
    _validate_deployment(config, path)
    return config
