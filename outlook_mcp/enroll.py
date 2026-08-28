"""Browser enrollment: how a user grants this server access to their mailbox.

An HTTP deployment only. A reverse proxy has already authenticated the visitor
and appended the user header, so the server knows who is asking without asking
anything itself; all that is missing is the visitor's consent at Microsoft.
These two routes collect it and write that user's MSAL cache.

Without them a person can only be enrolled by an operator running
``outlook-mcp-auth --user`` on the server, which does not scale past a handful
of mailboxes. The routes answer 404 unless [auth].public_url is set, because the
redirect URI handed to Entra has to be an absolute URL that Entra can send the
browser back to, and only the operator knows it.

Nothing here is reachable without the proxy: both routes run the same identity
policy the tools do, so the proxy must be configured to authenticate /oauth/*
exactly as it authenticates /mcp.
"""

import html
import logging
import secrets
import time
from typing import Dict, Optional, Tuple

import msal
from anyio import to_thread
from starlette.requests import Request
from starlette.responses import HTMLResponse, RedirectResponse, Response

from .app import get_config, mcp, proxy_auth_policy
from .auth import (
    GRAPH_SCOPE_URLS,
    CredentialsError,
    authority_for,
    save_token_cache,
    user_cache_path,
)
from .config import ENROLL_CALLBACK_PATH, ENROLL_LOGIN_PATH
from .credentials import Credentials, credentials_from_config

logger = logging.getLogger("outlook_mcp")

# How long a started sign-in may take to come back. Long enough for a password,
# an MFA prompt and a consent screen; short enough that an abandoned attempt
# does not sit in memory for the life of the process.
FLOW_TTL_SECONDS = 15 * 60

# Nothing legitimate needs many sign-ins in flight at once, and this dictionary
# is filled by anyone the proxy lets through, so it is bounded.
MAX_PENDING_FLOWS = 64

# state -> (flow, user, started_at). In memory on purpose: a restart losing a
# half-finished sign-in costs one click, while persisting it would put the PKCE
# verifier of an incomplete flow on disk.
_pending: Dict[str, Tuple[dict, str, float]] = {}


def _prune(now: float) -> None:
    for state, (_, _, started) in list(_pending.items()):
        if now - started > FLOW_TTL_SECONDS:
            del _pending[state]


def _remember(flow: dict, user: str) -> None:
    now = time.monotonic()
    _prune(now)
    if len(_pending) >= MAX_PENDING_FLOWS:
        # Drop the oldest rather than refusing: a full table would otherwise let
        # one abandoned browser tab lock everybody else out of enrolling.
        oldest = min(_pending, key=lambda s: _pending[s][2])
        del _pending[oldest]
    _pending[flow["state"]] = (flow, user, now)


def _take(state: str, user: str) -> Optional[dict]:
    """The flow this user started under `state`, consumed so it cannot replay."""
    entry = _pending.pop(state, None)
    if entry is None:
        return None
    flow, flow_user, started = entry
    if time.monotonic() - started > FLOW_TTL_SECONDS:
        return None
    # The proxy identity at the callback must be the one that started the flow.
    # Otherwise user A could finish user B's sign-in and have B's tokens filed
    # under A, which is the one mix-up this whole mode exists to prevent.
    if flow_user.strip().casefold() != user.strip().casefold():
        return None
    return flow


def _page(title: str, body: str, status: int = 200) -> HTMLResponse:
    """Render one of the enrollment pages.

    ``body`` is markup the caller assembles, so it is inserted as-is and every
    caller escapes its own values with ``_text()``. ``title`` is escaped here
    because it is plain text by contract.
    """
    safe_title = html.escape(title)
    return HTMLResponse(
        f"""<!doctype html>
<html><head><meta charset="utf-8"><title>{safe_title}</title></head>
<body style="font-family: system-ui; max-width: 34rem; margin: 4rem auto; line-height: 1.5">
<h1>{safe_title}</h1>
{body}
</body></html>""",
        status_code=status,
    )


def _text(value) -> str:
    """A value from outside the code, made safe to drop into markup.

    Three things reach these pages from elsewhere: the address the proxy
    asserted, the account name in the id_token, and an error description from
    Entra. None of them is markup, and the first is fully attacker-controlled
    the moment a proxy location forwards without replacing the identity header,
    which is exactly the misconfiguration SETUP.md warns about. Escaping keeps
    that mistake a failed sign-in rather than script execution on the origin
    that holds the proxy's session cookie.
    """
    return html.escape(str(value))


def _identify(request: Request) -> Tuple[str, Credentials]:
    """Who is asking, and with which app registration. Raises CredentialsError."""
    config = get_config()
    policy = proxy_auth_policy(config)
    if policy is None:  # pragma: no cover - the routes check enrollment_enabled first
        raise CredentialsError("Enrollment is only available over HTTP.")
    user = policy.user_from_headers(request.headers)
    creds = credentials_from_config(config)
    if creds is None:
        raise CredentialsError(
            "This server has no Azure AD app registration configured: set "
            "client_id and client_secret under [credentials] in outlook_mcp.toml "
            "and restart it."
        )
    return user, creds


def _msal_app(creds: Credentials, cache=None) -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        client_id=creds.client_id,
        client_credential=creds.client_secret,
        authority=authority_for(creds.tenant_id),
        token_cache=cache,
    )


@mcp.custom_route(ENROLL_LOGIN_PATH, methods=["GET"])
async def enroll_login(request: Request) -> Response:
    """Start the Microsoft sign-in that authorizes this server for one mailbox."""
    config = get_config()
    if not config.enrollment_enabled:
        return _page(
            "Not available",
            "<p>This server does not serve browser enrollment. Ask its operator "
            "to run <code>outlook-mcp-auth --user &lt;you&gt;</code>.</p>",
            status=404,
        )
    try:
        user, creds = _identify(request)
    except CredentialsError as e:
        return _page("Cannot identify you", f"<p>{_text(e)}</p>", status=403)

    app = _msal_app(creds)
    flow = await to_thread.run_sync(
        lambda: app.initiate_auth_code_flow(
            scopes=GRAPH_SCOPE_URLS,
            redirect_uri=config.enroll_callback_url,
            state=secrets.token_urlsafe(32),
        )
    )
    if "auth_uri" not in flow:
        logger.error("Could not start an authorization flow for %s: %s", user, flow)
        detail = flow.get("error_description", flow.get("error", "unknown error"))
        return _page("Cannot start sign-in", f"<p>{_text(detail)}</p>", status=502)

    _remember(flow, user)
    logger.info("Enrollment started for %s", user)
    return RedirectResponse(flow["auth_uri"], status_code=302)


@mcp.custom_route(ENROLL_CALLBACK_PATH, methods=["GET"])
async def enroll_callback(request: Request) -> Response:
    """Finish the sign-in and write this user's token cache."""
    config = get_config()
    if not config.enrollment_enabled:
        return _page("Not available", "<p>Enrollment is not served here.</p>", status=404)
    try:
        user, creds = _identify(request)
    except CredentialsError as e:
        return _page("Cannot identify you", f"<p>{_text(e)}</p>", status=403)

    params = dict(request.query_params)
    state = params.get("state", "")
    flow = _take(state, user)
    if flow is None:
        return _page(
            "Sign-in expired",
            f'<p>Start again from <a href="{ENROLL_LOGIN_PATH}">{ENROLL_LOGIN_PATH}</a>.</p>',
            status=400,
        )

    # A fresh cache, not the user's existing one: enrolling again replaces the
    # account rather than adding a second one to the same file, so the file
    # always holds exactly one and get_accounts() is never ambiguous.
    cache = msal.SerializableTokenCache()
    app = _msal_app(creds, cache)
    result = await to_thread.run_sync(
        lambda: app.acquire_token_by_auth_code_flow(flow, params)
    )

    if "access_token" not in result:
        # The description and the redirect_uri, not just the error code: this
        # step fails for reasons that all look like "invalid_request" from the
        # outside, and the operator reading this log is the one person who can
        # compare the URI against what the app registration holds. Nothing here
        # is a secret; a failed redemption returns no token.
        logger.warning(
            "Enrollment failed for %s: %s: %s (redirect_uri sent: %s)",
            user,
            result.get("error", "unknown"),
            result.get("error_description", "no description"),
            flow.get("redirect_uri", "unknown"),
        )
        detail = result.get("error_description", result.get("error", "unknown error"))
        return _page("Authorization failed", f"<p>{_text(detail)}</p>", status=400)

    path = user_cache_path(user, config.cache_directory)
    save_token_cache(cache, path)
    account = (result.get("id_token_claims") or {}).get("preferred_username", "")
    logger.info("Enrolled %s (Microsoft account %s)", user, account or "unknown")
    return _page(
        "Authorized",
        f"<p>This server can now reach the mailbox of "
        f"<strong>{_text(account or user)}</strong> on your behalf.</p>"
        f"<p>You can close this tab and use the Outlook tools.</p>",
    )
