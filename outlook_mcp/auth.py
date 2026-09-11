"""Authentication and Microsoft Graph API client."""

import hashlib
import logging
import os
from pathlib import Path
from typing import List, Optional, Tuple
from urllib.parse import urlparse

import httpx
import msal

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Mail, calendar and profile: what every grant has held since the first release.
CORE_SCOPES = [
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.Read",
    "Calendars.ReadWrite",
    "User.Read",
]
CONTACTS_SCOPES = ["Contacts.ReadWrite"]

# What a sign-in asks for: everything any tool needs. Both the authorization
# command and browser enrollment request exactly these, so they are defined once.
GRAPH_SCOPES = CORE_SCOPES + CONTACTS_SCOPES


def _scope_urls(scopes: List[str]) -> List[str]:
    """The scopes as Graph expects them on the wire."""
    return [f"https://graph.microsoft.com/{s}" for s in scopes]


GRAPH_SCOPE_URLS = _scope_urls(GRAPH_SCOPES)
CORE_SCOPE_URLS = _scope_urls(CORE_SCOPES)
CONTACTS_SCOPE_URLS = _scope_urls(CONTACTS_SCOPES)

# A request asks only for the scopes of the resource it addresses, never for the
# whole list. A refresh token redeems for the scopes its user has consented to
# and no more, and asking for a single one outside them fails the entire
# request (AADSTS70000 from a personal account, AADSTS65001 from a work one).
# Asking for everything would therefore have turned every mail and calendar tool
# off for each grant made before contacts existed, until its owner consented to
# the new scope. This way only the tools that need it ask for it.
_RESOURCE_SCOPE_URLS = {
    "/me/contacts": CONTACTS_SCOPE_URLS,
    "/me/contactfolders": CONTACTS_SCOPE_URLS,
}


def scopes_for(endpoint: str) -> List[str]:
    """The scope URLs a Graph request needs, from the resource it addresses.

    ``endpoint`` is a path relative to GRAPH_BASE_URL, or the absolute URL of an
    @odata.nextLink. Graph paths are case-insensitive, and the docs themselves
    spell both ``contactFolders`` and ``contactfolders``.
    """
    path = urlparse(endpoint).path.casefold()
    path = path.removeprefix(urlparse(GRAPH_BASE_URL).path)
    for prefix, scopes in _RESOURCE_SCOPE_URLS.items():
        if path.startswith(prefix):
            return scopes
    return CORE_SCOPE_URLS


# What AAD answers when a refresh token is asked for a scope its grant does not
# include: AADSTS70000 "one or more scopes requested are unauthorized or
# expired" (measured on a personal account), AADSTS65001 "has not consented"
# (a work account, which also sets the consent_required suberror).
_CONSENT_ERROR_CODES = {65001, 70000}


def _lacks_consent(result: Optional[dict]) -> bool:
    if not result or result.get("error") != "invalid_grant":
        return False
    codes = set(result.get("error_codes") or [])
    return bool(codes & _CONSENT_ERROR_CODES) or result.get("suberror") == "consent_required"

# Must match the redirect URI registered on the Azure AD app registration.
REDIRECT_URI = "http://localhost:5000/callback"

# The defaults, for a server nobody told where to keep its caches: one person's
# machine, where a dotted name under their home is the polite thing to be. A
# deployment that owns a directory of its own says so with [auth].cache_dir, and
# then nothing here is used: repeating the package name inside a directory that
# already belongs to this server would be noise, not namespacing.
TOKEN_CACHE_PATH = Path.home() / ".outlook_mcp_token_cache.json"

# One cache file per user, for the deployments where a reverse proxy says which
# user a request belongs to. Separate files rather than one shared cache: MSAL
# indexes its entries by client id, so a single app registration serving several
# people would put every account in one cache and get_accounts()[0] would return
# an arbitrary one. The isolation has to be the file, not a lookup key.
USER_CACHE_DIR = Path.home() / ".outlook_mcp" / "caches"

logger = logging.getLogger("outlook_mcp")


def authority_for(tenant_id: str) -> str:
    """The AAD authority URL for a tenant id (or "common")."""
    return f"https://login.microsoftonline.com/{tenant_id}"


def user_digest(user: str) -> str:
    """The filesystem name standing for one user, everywhere this server writes.

    A hash of the address rather than the address itself: it keeps the name
    portable whatever the address contains, and a directory listing then
    discloses how many people are served but not who they are. Whitespace and
    case are normalised first, because the proxy may spell the same person
    differently from one request to the next.
    """
    return hashlib.sha256(user.strip().casefold().encode("utf-8")).hexdigest()


def user_cache_path(user: str, directory: Optional[Path] = None) -> Path:
    """Where one user's MSAL cache lives.

    ``directory`` is [auth].cache_dir when a deployment set one. None means the
    default, so a caller can pass the configured value straight through without
    repeating the fallback.
    """
    return (directory or USER_CACHE_DIR) / f"{user_digest(user)}.json"


def shared_cache_path(directory: Optional[Path] = None) -> Path:
    """Where the cache for a single-account (stdio) server lives.

    Named rather than digested: there is no address to hash, because nothing in
    the request says who the caller is. One file, one account.
    """
    return TOKEN_CACHE_PATH if directory is None else directory / "shared.json"


class CredentialsError(RuntimeError):
    """Raised when a request cannot be served because no Azure AD credentials are available."""


def load_token_cache(path: Path = TOKEN_CACHE_PATH) -> msal.SerializableTokenCache:
    """Load the MSAL token cache written by outlook_mcp_auth.py (empty if absent)."""
    cache = msal.SerializableTokenCache()
    if path.exists():
        cache.deserialize(path.read_text())
    return cache


def _file_stamp(path: Path) -> Optional[Tuple[int, int]]:
    """What changes when a file is rewritten: its modification time and size.

    None when there is no file. Any other failure to stat it propagates: a
    cache that cannot be looked at cannot be written back either.
    """
    try:
        st = path.stat()
    except FileNotFoundError:
        return None
    return st.st_mtime_ns, st.st_size


def save_token_cache(cache: msal.SerializableTokenCache, path: Path) -> None:
    """Write a token cache, creating its directory and keeping it private.

    The file holds refresh tokens, so it is created 0600 before anything is
    written to it. On Windows the mode is ignored and the ACL inherited from
    the user profile is what protects it.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        # Create empty with the right mode first: writing then chmod-ing would
        # leave the tokens world-readable for the width of that window.
        os.close(os.open(path, os.O_CREAT | os.O_WRONLY, 0o600))
    path.write_text(cache.serialize())


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
        cache_path: Path = TOKEN_CACHE_PATH,
        user: Optional[str] = None,
    ):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = authority_for(tenant_id)
        # Set only when a reverse proxy told us whose mailbox this is. It marks
        # the manager as belonging to one enrolled person, which changes two
        # things: the errors talk about enrolling, and the app-only fallback is
        # off (see get_token), because acting as the application is precisely
        # what an unenrolled user must not be able to do.
        self.user = user
        # One manager, one file: the cache is read from the path it is written
        # back to. A manager holding one user's cache must never persist it over
        # the shared file, which is how every user would end up sharing one
        # account.
        self._cache_path = cache_path
        self._cache = load_token_cache(cache_path)
        # The file as this manager last read or wrote it. Anything else on disk
        # was written by someone else: see _adopt_rewritten_cache().
        self._cache_stamp = _file_stamp(cache_path)
        self._app: Optional[msal.ConfidentialClientApplication] = None
        # Whether AAD has confirmed this client secret at least once. Until it
        # has, no token may be served out of the cache. See get_token().
        self._secret_verified = False

    def _adopt_rewritten_cache(self) -> None:
        """Take up the cache file when someone else has rewritten it since.

        A sign-in writes it while the server runs: outlook-mcp-auth from another
        process, or /oauth/login in this one, to grant a new scope, to replace a
        grant that stopped working, or to put another account behind the file.
        Without this the server would go on with the copy it read at its first
        call until restarted, and its next write-back would put that copy back
        over the sign-in. A user who called a tool before enrolling would stay
        "not authorized" the same way. A deleted file is adopted too, as an
        empty cache: that is how an operator withdraws a grant.

        Reloaded in place, so the MSAL application keeps its cache object.
        """
        stamp = _file_stamp(self._cache_path)
        if stamp == self._cache_stamp:
            return
        self._cache.deserialize(self._cache_path.read_text() if stamp else "{}")
        self._cache_stamp = stamp
        logger.info("Token cache %s was rewritten, reloaded", self._cache_path.name)

    def _save_cache(self):
        """Persist the token cache, unless someone else has rewritten the file since.

        The other writer is a sign-in, and its grant is the newer one: this
        manager's tokens came from the grant it replaces. The next call adopts
        it rather than overwriting it.
        """
        if not self._cache.has_state_changed:
            return
        if _file_stamp(self._cache_path) != self._cache_stamp:
            return
        save_token_cache(self._cache, self._cache_path)
        self._cache_stamp = _file_stamp(self._cache_path)

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

    async def get_token(self, scopes: Optional[List[str]] = None) -> str:
        """Get a valid access token for `scopes`, refreshing if needed.

        `scopes` are the ones the request at hand needs (see scopes_for), the
        core mail, calendar and profile set when not given.

        The first token for a given set of credentials always costs one round
        trip to AAD, because that request is where AAD authenticates the client
        secret. MSAL keys cached tokens by client id alone (see the query built
        in acquire_token_silent), never by secret, so handing one out before the
        secret has been proven would let anyone who knows the client id, which
        is not secret material, use another caller's token. That is reachable
        only over HTTP, where the credentials arrive in request headers, but the
        guard belongs here where the token is produced.
        """
        scopes = scopes or CORE_SCOPE_URLS
        self._adopt_rewritten_cache()
        accounts = self.app.get_accounts()

        if accounts:
            # Redeeming the refresh token is a request AAD authenticates with
            # the client secret, so forcing it is what proves ownership. The
            # "with_error" variant, because acquire_token_silent turns every
            # refusal into None, and a missing consent needs its own answer.
            result = self.app.acquire_token_silent_with_error(
                scopes, account=accounts[0], force_refresh=not self._secret_verified
            )
            if result and "access_token" in result:
                self._secret_verified = True
                self._save_cache()
                return result["access_token"]
            if _lacks_consent(result):
                # Neither a wrong secret nor an expired grant: the grant is
                # fine for what it covers, and falling through to app-only
                # would be wrong for the same reason as for any delegated user.
                raise CredentialsError(self._consent_message(scopes, result))
            if not self._secret_verified:
                # A wrong secret and a stale refresh token look the same from
                # here, and we must not fall through to a cached token, so both
                # end as a credentials failure.
                raise CredentialsError(
                    "Could not obtain a token for this client id. Either the "
                    "client secret is wrong, or the cached authorization has "
                    f"expired: {self._reauthorize_hint()}."
                )

        if self.user is not None:
            # Per-user manager with nothing usable in its cache. Falling through
            # to client credentials would hand an unenrolled caller a token that
            # acts as the application itself, so the road stops here.
            raise CredentialsError(
                f"{self.user} has not authorized this server to reach their "
                f"mailbox, or that authorization has expired: {self._reauthorize_hint()}."
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

    def _reauthorize_hint(self) -> str:
        """How the operator of this particular manager fixes a missing grant."""
        if self.user is None:
            return "re-run python outlook_mcp_auth.py"
        return f"re-run outlook-mcp-auth --user {self.user}, or sign in again at /oauth/login"

    def _consent_message(self, scopes: List[str], result: dict) -> str:
        """Why a grant that works for other tools cannot serve this one."""
        names = ", ".join(s.rsplit("/", 1)[-1] for s in scopes)
        codes = ", ".join(f"AADSTS{c}" for c in result.get("error_codes") or []) or "invalid_grant"
        whose = f"for {self.user}" if self.user else "this server holds"
        return (
            f"The authorization {whose} does not include {names}, which this "
            f"tool needs ({codes}). "
            f"It was most likely granted before this server asked for it; the "
            f"other tools keep working meanwhile. To grant it, sign in once more "
            f"and accept the new permission: {self._reauthorize_hint()}."
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
        self, method: str, endpoint: str, headers: Optional[dict] = None, **kwargs
    ) -> dict:
        """Make an authenticated request to the Graph API.

        `headers` is added to the ones every request carries, e.g. a Prefer
        naming the time zone Graph should answer in. It cannot replace the
        Authorization header.

        `endpoint` may also be the absolute URL of an @odata.nextLink: httpx
        ignores base_url for an absolute URL, and the scopes are read off its
        path all the same.
        """
        token = await self.auth.get_token(scopes_for(endpoint))
        client = await self._get_client()
        headers = {
            "Content-Type": "application/json",
            **(headers or {}),
            "Authorization": f"Bearer {token}",
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

    async def get(
        self, endpoint: str, params: Optional[dict] = None, headers: Optional[dict] = None
    ) -> dict:
        return await self.request("GET", endpoint, params=params, headers=headers)

    async def post(self, endpoint: str, json_data: Optional[dict] = None) -> dict:
        return await self.request("POST", endpoint, json=json_data)

    async def patch(self, endpoint: str, json_data: Optional[dict] = None) -> dict:
        return await self.request("PATCH", endpoint, json=json_data)

    async def delete(self, endpoint: str) -> dict:
        return await self.request("DELETE", endpoint)
