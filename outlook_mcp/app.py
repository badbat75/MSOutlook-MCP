"""The shared MCPServer instance and its lifespan.

Tool modules import ``mcp`` from here rather than from server.py: the server
object has to exist before any ``@mcp.tool()`` decorator runs, while server.py
is the entry point that imports the tools. Keeping the object here is what
breaks that circle.
"""

import logging
from contextlib import asynccontextmanager

from mcp.server.mcpserver import MCPServer

from .config import ServerConfig
from .credentials import (
    ENV_CLIENT_ID,
    ENV_CLIENT_SECRET,
    HEADER_CLIENT_ID,
    HEADER_CLIENT_SECRET,
    HEADER_TENANT_ID,
    GraphClientPool,
    credentials_from_env,
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


@asynccontextmanager
async def app_lifespan(app):
    """Create the Graph client pool on startup, close every client on shutdown."""
    pool = GraphClientPool()
    env_creds = credentials_from_env()

    if _config.is_http:
        logger.info(
            "HTTP mode: credentials are taken from the %s / %s / %s request headers",
            HEADER_CLIENT_ID, HEADER_CLIENT_SECRET, HEADER_TENANT_ID,
        )
    elif env_creds is None:
        logger.warning(
            "%s and %s must be set. The server will start but all tools will fail "
            "until configured.", ENV_CLIENT_ID, ENV_CLIENT_SECRET,
        )

    yield {"pool": pool, "env_credentials": env_creds}

    await pool.close_all()


mcp = MCPServer("MS_Outlook_MCP", lifespan=app_lifespan)
