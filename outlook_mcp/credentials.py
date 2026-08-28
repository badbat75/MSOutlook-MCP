"""Where a request's Azure AD credentials come from, and the clients they own.

The transport decides the source, not a flag. Over HTTP the SDK attaches the
Starlette request to the tool context and the ``X-Outlook-*`` headers of that
request are authoritative; over stdio there is no request object and the
``OUTLOOK_*`` variables read at startup are used.
"""

import logging
import os
from dataclasses import dataclass
from typing import Dict, Mapping, Optional

from mcp.server.mcpserver import Context

from .auth import AuthManager, CredentialsError, GraphClient, load_token_cache

logger = logging.getLogger("outlook_mcp")

HEADER_CLIENT_ID = "X-Outlook-Client-Id"
HEADER_CLIENT_SECRET = "X-Outlook-Client-Secret"
HEADER_TENANT_ID = "X-Outlook-Tenant-Id"

ENV_CLIENT_ID = "OUTLOOK_CLIENT_ID"
ENV_CLIENT_SECRET = "OUTLOOK_CLIENT_SECRET"
ENV_TENANT_ID = "OUTLOOK_TENANT_ID"

DEFAULT_TENANT_ID = "common"


@dataclass(frozen=True)
class Credentials:
    """Azure AD app registration credentials for one Graph client."""

    client_id: str
    client_secret: str
    tenant_id: str = DEFAULT_TENANT_ID


def credentials_from_env(environ: Mapping[str, str] = os.environ) -> Optional[Credentials]:
    """Read credentials from OUTLOOK_* variables; None when incomplete."""
    client_id = environ.get(ENV_CLIENT_ID, "").strip()
    client_secret = environ.get(ENV_CLIENT_SECRET, "").strip()
    if not client_id or not client_secret:
        return None
    tenant_id = environ.get(ENV_TENANT_ID, "").strip() or DEFAULT_TENANT_ID
    return Credentials(client_id, client_secret, tenant_id)


def credentials_from_headers(headers: Mapping[str, str]) -> Credentials:
    """Read credentials from the X-Outlook-* request headers (case-insensitive).

    Raises CredentialsError naming the missing headers, so a misconfigured
    client sees exactly what to send instead of a generic auth failure.
    """
    # Starlette's Headers are already case-insensitive; normalising here keeps
    # the function correct for any plain Mapping too (tests, other frameworks).
    lowered = {str(k).lower(): v for k, v in headers.items()}
    client_id = (lowered.get(HEADER_CLIENT_ID.lower()) or "").strip()
    client_secret = (lowered.get(HEADER_CLIENT_SECRET.lower()) or "").strip()
    missing = [
        name for name, value in (
            (HEADER_CLIENT_ID, client_id),
            (HEADER_CLIENT_SECRET, client_secret),
        ) if not value
    ]
    if missing:
        raise CredentialsError(
            f"Missing required HTTP header(s): {', '.join(missing)}. In HTTP mode the "
            f"Azure AD credentials must be sent with every request as "
            f"{HEADER_CLIENT_ID}, {HEADER_CLIENT_SECRET} and (optionally) "
            f"{HEADER_TENANT_ID}; environment variables are not consulted."
        )
    tenant_id = (lowered.get(HEADER_TENANT_ID.lower()) or "").strip() or DEFAULT_TENANT_ID
    return Credentials(client_id, client_secret, tenant_id)


class GraphClientPool:
    """One GraphClient per distinct set of credentials, created on first use.

    In stdio mode the pool only ever holds the environment credentials. In HTTP
    mode each caller supplies its own headers, so the pool keeps a client per
    app registration and all of them share one MSAL token cache (tokens are
    keyed by client id inside the cache, and one serialize() persists them all).
    """

    def __init__(self):
        self._token_cache = load_token_cache()
        self._clients: Dict[Credentials, GraphClient] = {}

    def get(self, creds: Credentials) -> GraphClient:
        client = self._clients.get(creds)
        if client is None:
            auth = AuthManager(
                creds.client_id, creds.client_secret, creds.tenant_id,
                token_cache=self._token_cache,
            )
            client = GraphClient(auth)
            self._clients[creds] = client
            logger.info(
                "Created Graph client for client_id=%s... tenant=%s",
                creds.client_id[:8], creds.tenant_id,
            )
        return client

    async def close_all(self):
        for client in self._clients.values():
            await client.close()
        self._clients.clear()


def get_graph(ctx: Context) -> GraphClient:
    """Resolve the GraphClient for the current request.

    The transport decides where credentials come from: over HTTP the SDK
    attaches the Starlette request to the context and the X-Outlook-* headers
    of that request are authoritative; over stdio there is no request object
    and the OUTLOOK_* environment variables read at startup are used.
    """
    request_context = ctx.request_context
    pool: GraphClientPool = request_context.lifespan_context["pool"]
    request = request_context.request
    if request is not None and hasattr(request, "headers"):
        creds = credentials_from_headers(request.headers)
    else:
        creds = request_context.lifespan_context["env_credentials"]
        if creds is None:
            raise CredentialsError(
                f"{ENV_CLIENT_ID} and {ENV_CLIENT_SECRET} are not set. Configure them in "
                f"the environment of the MCP server (e.g. the \"env\" block of the Claude "
                f"Desktop config) and restart the server."
            )
    return pool.get(creds)
