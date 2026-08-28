"""The shared MCPServer instance and its lifespan.

Tool modules import ``mcp`` from here rather than from server.py: the server
object has to exist before any ``@mcp.tool()`` decorator runs, while server.py
is the entry point that imports the tools. Keeping the object here is what
breaks that circle.
"""

import logging
from contextlib import asynccontextmanager
from typing import Optional

import anyio
from mcp.server.mcpserver import MCPServer

from .config import CONFIG_FILENAME, ServerConfig
from .credentials import (
    GraphClientPool,
    ProxyAuthPolicy,
    credentials_from_config,
)

logger = logging.getLogger("outlook_mcp")

# Listening parameters come from the external config file (see config.py).
# main() replaces this with the loaded file; the default keeps `mcp dev` and
# direct imports working as a plain stdio server.
_config = ServerConfig()


def get_config() -> ServerConfig:
    """The configuration the server is running with."""
    return _config


def set_config(config: ServerConfig) -> None:
    """Install the configuration loaded by the entry point. Call before run()."""
    global _config
    _config = config


def proxy_auth_policy(config: ServerConfig) -> Optional[ProxyAuthPolicy]:
    """The identity policy for HTTP requests, or None over stdio."""
    if not config.is_http:
        return None
    return ProxyAuthPolicy(user_header=config.user_header)


@asynccontextmanager
async def app_lifespan(app):
    """Create the Graph client pool on startup, close every client on shutdown.

    It also owns the download reaper, because a background task has to be
    cancelled by whatever started it and this is the one scope that spans the
    life of the server.
    """
    # Imported here and not at module scope: downloads.py registers its route on
    # `mcp`, so it imports this module and cannot be imported by it.
    from .downloads import reap_expired_downloads

    pool = GraphClientPool(_config.cache_directory)
    credentials = credentials_from_config(_config)
    proxy_auth = proxy_auth_policy(_config)

    if proxy_auth is not None:
        logger.info(
            "HTTP mode: the mailbox is chosen by the %s header of each request, "
            "served with the configured app registration and one token cache per "
            "user. Listening on %s only.",
            proxy_auth.user_header, _config.bind_host,
        )
    if credentials is None:
        logger.warning(
            "No [credentials] in %s: the server will start but every tool will "
            "fail until client_id and client_secret are configured.",
            _config.source or CONFIG_FILENAME,
        )

    async with anyio.create_task_group() as housekeeping:
        if _config.retention_seconds is not None:
            logger.info(
                "Retention for downloads nobody fetches: %d minutes.",
                _config.retention_minutes,
            )
            housekeeping.start_soon(reap_expired_downloads)
        try:
            yield {"pool": pool, "credentials": credentials, "proxy_auth": proxy_auth}
        finally:
            # The reaper loops forever by design, so shutdown is a cancellation
            # rather than something to wait for.
            housekeeping.cancel_scope.cancel()

    await pool.close_all()


mcp = MCPServer("MS_Outlook_MCP", lifespan=app_lifespan)
