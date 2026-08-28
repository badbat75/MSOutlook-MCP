"""Whose mailbox a request acts on, and the Graph clients that serve it.

The server has exactly one Azure AD app registration, read from its
configuration file, and it never accepts one from a caller. What varies per
request is only *whose* mailbox is being addressed:

* Over **stdio** there is no request and no ambiguity. The process was started
  by one person, on their own machine or under their own account, and the
  operating system user is the whole authentication story: whoever can spawn
  the server can already read its configuration and its token cache.

* Over **HTTP** a reverse proxy authenticates the user and appends the user
  header itself. The header carries no proof, and needs none, because the
  server listens on loopback only and the proxy is therefore the sole route in,
  which config.py enforces at startup. Each user gets a token cache of their
  own, so one person's grant is never served to another.
"""

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Mapping, Optional

from mcp.server.mcpserver import Context

from .auth import (
    AuthManager,
    CredentialsError,
    GraphClient,
    load_token_cache,
    shared_cache_path,
    user_cache_path,
)

logger = logging.getLogger("outlook_mcp")


@dataclass(frozen=True)
class Credentials:
    """Azure AD app registration credentials for one Graph client."""

    client_id: str
    client_secret: str
    tenant_id: str


def credentials_from_config(config) -> Optional[Credentials]:
    """The app registration the server runs as, or None when unconfigured."""
    if not config.has_credentials:
        return None
    return Credentials(config.client_id, config.client_secret, config.tenant_id)


@dataclass(frozen=True)
class ProxyAuthPolicy:
    """The identity a reverse proxy asserts, and where to read it.

    There is deliberately nothing here that authenticates the proxy: the header
    is believed because the server listens on loopback only and the proxy is
    therefore the sole route in, which config.py enforces at startup. Putting a
    second check here would suggest the mode is safe without that, and it is not.
    """

    user_header: str

    def user_from_headers(self, headers: Mapping[str, str]) -> str:
        """The address the proxy vouched for, or CredentialsError."""
        # Starlette's Headers are already case-insensitive; normalising here
        # keeps the method correct for a plain Mapping too (tests, other
        # frameworks), and for a proxy that spells the header its own way.
        lowered = {str(k).lower(): v for k, v in headers.items()}
        user = (lowered.get(self.user_header.lower()) or "").strip()
        if not user:
            raise CredentialsError(
                f"No {self.user_header} header on this request. The reverse proxy "
                f"in front of this server sets it once it has authenticated the "
                f"user; a request arriving without it cannot be attributed to a "
                f"mailbox."
            )
        return user


@dataclass(frozen=True)
class Principal:
    """Whose mailbox a request acts on, and with which app registration.

    ``user`` is set only over HTTP, where one app registration serves everyone
    and the enrolled person is what distinguishes one caller from the next. Over
    stdio there is only ever one person and it stays None. The pair is the
    pool's key, so two people are never one client.
    """

    credentials: Credentials
    user: Optional[str] = None


class GraphClientPool:
    """One GraphClient per principal, created on first use.

    Over stdio the pool only ever holds one entry, backed by the single token
    cache the authorization command writes. Over HTTP the app registration is
    the same for everyone, so sharing one cache would be a leak: MSAL indexes
    its entries by client id, every account would land in the same cache, and
    get_accounts() would return an arbitrary one. Each user therefore gets a
    cache object of its own, loaded from and written back to its own file.
    """

    def __init__(self, cache_dir: Optional[Path] = None):
        # None means the home directory default, which is what auth.py applies.
        self._cache_dir = cache_dir
        self._shared_path = shared_cache_path(cache_dir)
        self._token_cache = load_token_cache(self._shared_path)
        self._clients: Dict[Principal, GraphClient] = {}

    def get(self, principal: Principal) -> GraphClient:
        client = self._clients.get(principal)
        if client is None:
            creds = principal.credentials
            if principal.user is None:
                cache, cache_path = self._token_cache, self._shared_path
            else:
                cache_path = user_cache_path(principal.user, self._cache_dir)
                cache = load_token_cache(cache_path)
            auth = AuthManager(
                creds.client_id, creds.client_secret, creds.tenant_id,
                token_cache=cache,
                cache_path=cache_path,
                user=principal.user,
            )
            client = GraphClient(auth)
            self._clients[principal] = client
            logger.info(
                "Created Graph client for client_id=%s... tenant=%s user=%s",
                creds.client_id[:8], creds.tenant_id, principal.user or "-",
            )
        return client

    async def close_all(self):
        for client in self._clients.values():
            await client.close()
        self._clients.clear()


def current_user(ctx: Context) -> Optional[str]:
    """Whose mailbox the current request speaks for, or None over stdio.

    The transport answers it: over stdio there is no request object and no
    ambiguity, over HTTP the SDK attaches the Starlette request and the identity
    header the reverse proxy appended names the user. Tools that write or delete
    files on the server ask this too, so that what belongs to one caller is
    never handed to, or removed by, another.
    """
    request_context = ctx.request_context
    proxy_auth: Optional[ProxyAuthPolicy] = request_context.lifespan_context.get("proxy_auth")
    request = request_context.request

    if request is None or not hasattr(request, "headers") or proxy_auth is None:
        return None
    return proxy_auth.user_from_headers(request.headers)


def get_graph(ctx: Context) -> GraphClient:
    """Resolve the GraphClient for the current request."""
    lifespan_context = ctx.request_context.lifespan_context

    credentials: Optional[Credentials] = lifespan_context["credentials"]
    if credentials is None:
        raise CredentialsError(
            "This server has no Azure AD app registration configured. Set "
            "client_id and client_secret under [credentials] in outlook_mcp.toml "
            "and restart it."
        )

    pool: GraphClientPool = lifespan_context["pool"]
    return pool.get(Principal(credentials, current_user(ctx)))
