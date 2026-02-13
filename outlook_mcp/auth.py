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

TOKEN_CACHE_PATH = Path.home() / ".outlook_mcp_token_cache.json"

logger = logging.getLogger("outlook_mcp")


# =============================================================================
# Authentication Manager
# =============================================================================

class AuthManager:
    """Handles MSAL authentication with token caching and refresh."""

    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._cache = msal.SerializableTokenCache()
        self._app: Optional[msal.ConfidentialClientApplication] = None
        self._load_cache()

    def _load_cache(self):
        """Load token cache from disk."""
        if TOKEN_CACHE_PATH.exists():
            self._cache.deserialize(TOKEN_CACHE_PATH.read_text())

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

    def get_auth_url(self) -> str:
        """Generate the authorization URL for initial user consent."""
        scopes = [f"https://graph.microsoft.com/{s}" for s in GRAPH_SCOPES]
        flow = self.app.initiate_auth_code_flow(
            scopes=scopes,
            redirect_uri="http://localhost:5000/callback",
        )
        self._pending_flow = flow
        return flow["auth_uri"]

    def complete_auth(self, auth_response: dict) -> dict:
        """Complete the auth flow with the callback response."""
        result = self.app.acquire_token_by_auth_code_flow(
            self._pending_flow, auth_response
        )
        self._save_cache()
        return result

    async def get_token(self) -> str:
        """Get a valid access token, refreshing if needed."""
        scopes = [f"https://graph.microsoft.com/{s}" for s in GRAPH_SCOPES]
        accounts = self.app.get_accounts()

        if accounts:
            result = self.app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache()
                return result["access_token"]

        # If no cached token, try client credentials (app-only)
        result = self.app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if result and "access_token" in result:
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
        response.raise_for_status()
        if response.status_code == 204:
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
