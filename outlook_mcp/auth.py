"""Authentication and Microsoft Graph API client."""

import logging
from pathlib import Path
from typing import Optional

import httpx
import msal

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = [
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.Read",
    "Calendars.ReadWrite",
    "User.Read",
]

# The same list as Graph expects it on the wire. Both the server and the
# authorization flow request exactly these, so they are defined once.
GRAPH_SCOPE_URLS = [f"https://graph.microsoft.com/{s}" for s in GRAPH_SCOPES]

# Must match the redirect URI registered on the Azure AD app registration.
REDIRECT_URI = "http://localhost:5000/callback"

TOKEN_CACHE_PATH = Path.home() / ".outlook_mcp_token_cache.json"

logger = logging.getLogger("outlook_mcp")


def authority_for(tenant_id: str) -> str:
    """The AAD authority URL for a tenant id (or "common")."""
    return f"https://login.microsoftonline.com/{tenant_id}"


class CredentialsError(RuntimeError):
    """Raised when a request cannot be served because no Azure AD credentials are available."""


def load_token_cache(path: Path = TOKEN_CACHE_PATH) -> msal.SerializableTokenCache:
    """Load the MSAL token cache written by outlook_mcp_auth.py (empty if absent).

    One cache instance can back several ``AuthManager`` objects: MSAL keys every
    entry by client id, so tokens for different app registrations coexist and a
    single serialize() call persists all of them. Give each AuthManager its own
    cache object only if they must never see each other's tokens.
    """
    cache = msal.SerializableTokenCache()
    if path.exists():
        cache.deserialize(path.read_text())
    return cache


# =============================================================================
# Authentication Manager
# =============================================================================

class AuthManager:
    """Handles MSAL authentication with token caching and refresh."""

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        tenant_id: str,
        token_cache: Optional[msal.SerializableTokenCache] = None,
    ):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = authority_for(tenant_id)
        # A caller may hand in a cache shared with other managers (HTTP mode,
        # one manager per set of header credentials); otherwise load our own.
        self._cache = token_cache if token_cache is not None else load_token_cache()
        self._app: Optional[msal.ConfidentialClientApplication] = None
        # Whether AAD has confirmed this client secret at least once. Until it
        # has, no token may be served out of the cache. See get_token().
        self._secret_verified = False

    def _save_cache(self):
        """Persist token cache to disk."""
        if self._cache.has_state_changed:
            TOKEN_CACHE_PATH.write_text(self._cache.serialize())

    @property
    def app(self) -> msal.ConfidentialClientApplication:
        if self._app is None:
            self._app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=self.authority,
                token_cache=self._cache,
            )
        return self._app

    def _unverified_client_app(self) -> msal.ConfidentialClientApplication:
        """An app with a private, empty cache, so its calls always reach AAD.

        Used for the app-only path before the secret has been proven: MSAL
        refuses force_refresh on acquire_token_for_client, and a cache that
        holds an app token would answer it locally without ever contacting AAD.
        """
        return msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority,
        )

    async def get_token(self) -> str:
        """Get a valid access token, refreshing if needed.

        The first token for a given set of credentials always costs one round
        trip to AAD, because that request is where AAD authenticates the client
        secret. MSAL keys cached tokens by client id alone (see the query built
        in acquire_token_silent), never by secret, so handing one out before the
        secret has been proven would let anyone who knows the client id, which
        is not secret material, use another caller's token. That is reachable
        only over HTTP, where the credentials arrive in request headers, but the
        guard belongs here where the token is produced.
        """
        scopes = GRAPH_SCOPE_URLS
        accounts = self.app.get_accounts()

        if accounts:
            # Redeeming the refresh token is a request AAD authenticates with
            # the client secret, so forcing it is what proves ownership.
            result = self.app.acquire_token_silent(
                scopes, account=accounts[0], force_refresh=not self._secret_verified
            )
            if result and "access_token" in result:
                self._secret_verified = True
                self._save_cache()
                return result["access_token"]
            if not self._secret_verified:
                # A wrong secret and a stale refresh token look the same from
                # here, and we must not fall through to a cached token, so both
                # end as a credentials failure.
                raise CredentialsError(
                    "Could not obtain a token for this client id. Either the "
                    "client secret is wrong, or the cached authorization has "
                    "expired: re-run python outlook_mcp_auth.py."
                )

        # No delegated account for this client id: client credentials (app-only).
        app = self.app if self._secret_verified else self._unverified_client_app()
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if result and "access_token" in result:
            self._secret_verified = True
            self._save_cache()
            return result["access_token"]

        raise RuntimeError(
            "No valid token available. Run the auth setup script first: "
            "python outlook_mcp_auth.py"
        )


# =============================================================================
# Microsoft Graph API Client
# =============================================================================

class GraphClient:
    """Async HTTP client for Microsoft Graph API."""

    def __init__(self, auth_manager: AuthManager):
        self.auth = auth_manager
        self._client: Optional[httpx.AsyncClient] = None

    async def _get_client(self) -> httpx.AsyncClient:
        if self._client is None or self._client.is_closed:
            self._client = httpx.AsyncClient(
                base_url=GRAPH_BASE_URL,
                timeout=30.0,
            )
        return self._client

    async def close(self):
        if self._client and not self._client.is_closed:
            await self._client.aclose()

    async def request(
        self, method: str, endpoint: str, **kwargs
    ) -> dict:
        """Make an authenticated request to the Graph API."""
        token = await self.auth.get_token()
        client = await self._get_client()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        response = await client.request(
            method, endpoint, headers=headers, **kwargs
        )
        if response.status_code >= 400:
            # Log the outgoing payload and Graph's error body so 4xx/5xx causes
            # are diagnosable (httpx only logs the request line, not the body).
            logger.error(
                "Graph %s %s -> %s | params=%s | body=%s | response=%s",
                method,
                endpoint,
                response.status_code,
                kwargs.get("params"),
                kwargs.get("json"),
                response.text,
            )
        response.raise_for_status()
        # Some Graph endpoints return an empty body: sendMail → 202 Accepted,
        # delete/update → 204 No Content. Don't try to JSON-decode those.
        if response.status_code in (202, 204) or not response.content:
            return {"status": "success"}
        return response.json()

    async def get(self, endpoint: str, params: Optional[dict] = None) -> dict:
        return await self.request("GET", endpoint, params=params)

    async def post(self, endpoint: str, json_data: Optional[dict] = None) -> dict:
        return await self.request("POST", endpoint, json=json_data)

    async def patch(self, endpoint: str, json_data: Optional[dict] = None) -> dict:
        return await self.request("PATCH", endpoint, json=json_data)

    async def delete(self, endpoint: str) -> dict:
        return await self.request("DELETE", endpoint)
