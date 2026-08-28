"""
Outlook MCP - OAuth2 Authentication Setup
==========================================
Run once to authorize the MCP server to access an Outlook account. Exposed as
the ``outlook-mcp-auth`` command and as the outlook_mcp_auth.py wrapper.

Three authorization modes:

1. Normal mode (default) - Opens browser, waits for callback, allows Ctrl+C fallback
   python outlook_mcp_auth.py

2. Headless mode - For remote/SSH systems, prompts for manual URL input
   python outlook_mcp_auth.py --no-browser

3. Direct mode - Provide authorization code or callback URL directly
   python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'

4. Per user, for an HTTP deployment where a reverse proxy names the user
   outlook-mcp-auth --user someone@example.com

The app registration comes from the same outlook_mcp.toml the server reads, so
authorizing needs no separate setup step.

This is the only place the authorization code flow lives. It talks to MSAL
directly rather than through AuthManager: that class serves an already
authorized account, and its token acquisition now deliberately forces a
refresh, which is the wrong behaviour for a first sign-in.
"""

import argparse
import json
import sys
import webbrowser
from html import escape
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import msal

from .auth import (
    GRAPH_SCOPE_URLS,
    REDIRECT_URI,
    authority_for,
    load_token_cache,
    save_token_cache,
    shared_cache_path,
    user_cache_path,
)
from .config import CONFIG_FILENAME, ConfigError, load_config


class CallbackHandler(BaseHTTPRequestHandler):
    """HTTP handler to capture the OAuth callback."""

    auth_code = None
    full_url = None

    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)

        if "code" in params:
            CallbackHandler.auth_code = params["code"][0]
            CallbackHandler.full_url = self.path
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            self.wfile.write(b"""
                <html><body style="font-family: system-ui; text-align: center; margin-top: 100px;">
                <h1>&#10003; Authorization Successful!</h1>
                <p>You can close this window and return to the terminal.</p>
                </body></html>
            """)
        elif "error" in params:
            # Escaped because it is a query string being echoed back, even
            # though this listener answers on loopback for a few seconds and
            # only to the person who started it.
            error = escape(params.get("error", ["unknown"])[0])
            desc = escape(params.get("error_description", [""])[0])
            self.send_response(400)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            self.wfile.write(f"""
                <html><body style="font-family: system-ui; text-align: center; margin-top: 100px;">
                <h1>&#10007; Authorization Failed</h1>
                <p>Error: {error}</p>
                <p>{desc}</p>
                </body></html>
            """.encode())
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        pass  # Suppress default logging


def _print_credentials_help(source) -> None:
    """Explain how to get the credentials this command needs."""
    print("=" * 60)
    print("ERROR: Azure AD credentials not set!")
    print("=" * 60)
    print()
    if source:
        print(f"Read {source}, but it has no [credentials] table.")
    else:
        print(f"No {CONFIG_FILENAME} found. Copy {CONFIG_FILENAME}.example to")
        print(f"{CONFIG_FILENAME} in the project root and fill it in.")
    print()
    print("  [credentials]")
    print("  client_id = 'your-client-id'")
    print("  client_secret = 'your-client-secret'")
    print("  tenant_id = 'common'")
    print()
    print("To get these values:")
    print("  1. Go to https://entra.microsoft.com")
    print("  2. Navigate to: Identity > Applications > App registrations")
    print("  3. Click 'New registration'")
    print("  4. Name: 'Outlook MCP Server'")
    print("  5. Supported account types: pick your preference")
    print(f"  6. Redirect URI: Web > {REDIRECT_URI}")
    print("  7. After creation, copy the Application (client) ID")
    print("  8. Go to 'Certificates & secrets' > New client secret")
    print("  9. Go to 'API permissions' > Add permission > Microsoft Graph:")
    for scope in GRAPH_SCOPE_URLS:
        name = scope.split("/")[-1]
        print(f"     - {name} (Delegated)")
    print("  10. Click 'Grant admin consent' (if applicable)")
    print()


def _auth_response_from(value: str) -> dict:
    """Turn a pasted callback URL, or a bare code, into MSAL's auth response."""
    if value.startswith("http"):
        parsed = urlparse(value)
        return {k: v[0] for k, v in parse_qs(parsed.query).items()}
    return {"code": value}


def _survive_a_narrow_console() -> None:
    """Never let a character this console cannot render abort a sign-in.

    A Windows console is cp1252, and printing anything outside it raises
    UnicodeEncodeError mid-flow: an arrow in the setup instructions used to kill
    the run after the browser had already opened. The output here is plain ASCII
    for that reason, and this makes a future slip degrade to '?' instead of
    losing the authorization.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")


def main():
    _survive_a_narrow_console()

    parser = argparse.ArgumentParser(description="Outlook MCP OAuth2 Setup")
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Don't automatically open browser (for headless/remote systems)",
    )
    parser.add_argument(
        "--code",
        type=str,
        help="Manually provide the authorization code or full callback URL",
    )
    parser.add_argument(
        "--user",
        metavar="EMAIL",
        help=(
            "Authorize on behalf of one user of an HTTP deployment, writing that "
            "user's own token cache instead of the shared one. The address is the "
            "one the reverse proxy puts in its user header."
        ),
    )
    parser.add_argument(
        "--config",
        metavar="PATH",
        help=(
            "Path to the TOML configuration file holding [credentials]. Defaults "
            "to $OUTLOOK_MCP_CONFIG, then outlook_mcp.toml in the project root."
        ),
    )
    args = parser.parse_args()

    # The same configuration file the server reads, so authorizing needs no
    # separate setup step. It has to come first: it is also what says where the
    # caches live, so writing one before reading it would put this run's tokens
    # somewhere the server never looks.
    try:
        config = load_config(args.config)
    except ConfigError as e:
        print(f"Configuration error: {e}", file=sys.stderr)
        sys.exit(2)

    # Which cache this run fills. The shared file is the stdio default; --user is
    # the multi-user deployment, where every mailbox gets its own.
    cache_dir = config.cache_directory
    cache_path = (
        user_cache_path(args.user, cache_dir) if args.user
        else shared_cache_path(cache_dir)
    )

    if not config.has_credentials:
        _print_credentials_help(config.source)
        sys.exit(1)

    client_id = config.client_id
    client_secret = config.client_secret
    tenant_id = config.tenant_id

    def build_app(token_cache):
        return msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority_for(tenant_id),
            token_cache=token_cache,
        )

    # Initialize MSAL
    cache = load_token_cache(cache_path)
    app = build_app(cache)

    # Check if we already have a valid token
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPE_URLS, account=accounts[0])
        if result and "access_token" in result:
            print("OK: already authenticated, the token is still valid.")
            print(f"   Account: {accounts[0].get('username', 'unknown')}")
            print(f"   Token cache: {cache_path}")
            if cache.has_state_changed:
                save_token_cache(cache, cache_path)
            return

    if args.user:
        # One file, one account. A per-user cache that accumulated two accounts
        # would leave get_accounts() picking between them arbitrarily, so a new
        # sign-in for this user replaces the previous grant instead of joining
        # it. The shared cache is different: it legitimately holds one account
        # per app registration, keyed by client id.
        cache = msal.SerializableTokenCache()
        app = build_app(cache)

    # Start auth code flow
    flow = app.initiate_auth_code_flow(
        scopes=GRAPH_SCOPE_URLS,
        redirect_uri=REDIRECT_URI,
    )

    if "auth_uri" not in flow:
        print("ERROR: Failed to create authorization URL.")
        print(json.dumps(flow, indent=2))
        sys.exit(1)

    auth_url = flow["auth_uri"]
    print("=" * 60)
    print("OUTLOOK MCP - OAuth2 Setup")
    print("=" * 60)
    print()

    if args.no_browser:
        print("MANUAL AUTHORIZATION REQUIRED")
        print()
        print("Copy this URL and open it on ANY device with a browser:")
        print()
        print(f"  {auth_url}")
        print()
        print("Steps:")
        print("  1. Copy the URL above")
        print("  2. Open it in a browser on any device (phone, laptop, etc.)")
        print("  3. Sign in with your Microsoft account")
        print("  4. After sign-in, browser will redirect to localhost:5000/callback")
        print("  5. Copy the FULL redirect URL from your browser address bar")
        print(f"     (It will look like: {REDIRECT_URI}?code=...)")
        print("  6. Paste it below when prompted")
        print()
    else:
        print("Opening browser for Microsoft login...")
        print(f"If browser doesn't open, visit:\n{auth_url}")
        print()
        webbrowser.open(auth_url)

    # Get authorization code
    if args.code:
        # Manual mode: user provides code or full URL via command line
        print()
        print("Using manually provided authorization code/URL...")
        auth_response = _auth_response_from(args.code)

    elif args.no_browser:
        # Headless mode: skip callback server, go straight to manual input
        print("=" * 60)
        print("Paste Callback URL")
        print("=" * 60)
        print()
        callback_url = input("Callback URL: ").strip()

        if not callback_url:
            print()
            print("ERROR: no URL provided. Exiting.")
            sys.exit(1)
        auth_response = _auth_response_from(callback_url)

    else:
        # Normal mode: callback server with Ctrl+C fallback
        print("Waiting for authorization callback on http://localhost:5000 ...")
        print()
        print("TIP: if the callback doesn't work, press Ctrl+C to paste manually")
        print()

        server = HTTPServer(("localhost", 5000), CallbackHandler)

        try:
            while CallbackHandler.auth_code is None:
                server.handle_request()

            server.server_close()

            # Got callback, parse it
            parsed = urlparse(f"http://localhost:5000{CallbackHandler.full_url}")
            auth_response = {k: v[0] for k, v in parse_qs(parsed.query).items()}

        except KeyboardInterrupt:
            server.server_close()
            print()
            print()
            print("=" * 60)
            print("Manual Authorization")
            print("=" * 60)
            print()
            print("Paste the FULL callback URL from your browser:")
            print(f"(It should look like: {REDIRECT_URI}?code=...)")
            print()

            callback_url = input("Callback URL: ").strip()

            if not callback_url:
                print()
                print("ERROR: no URL provided. Exiting.")
                sys.exit(1)
            auth_response = _auth_response_from(callback_url)

    result = app.acquire_token_by_auth_code_flow(flow, auth_response)

    if "access_token" in result:
        # Save cache
        save_token_cache(cache, cache_path)

        print()
        print("OK: authentication successful.")
        if args.user:
            print(f"   Authorized for: {args.user}")
        print(f"   Token cache saved to: {cache_path}")
        # MSAL hands back whatever AAD sent: a space delimited string here, a
        # list in other flows. Joining the string would print it one character
        # at a time.
        granted = result.get("scope") or []
        if isinstance(granted, str):
            granted = granted.split()
        print(f"   Scopes granted: {', '.join(granted)}")
        print()
        print("You can now start the MCP server:")
        print("   python outlook_mcp_server.py")
    else:
        print()
        print("ERROR: authentication failed.")
        print(f"   Error: {result.get('error', 'unknown')}")
        print(f"   Description: {result.get('error_description', 'N/A')}")
        sys.exit(1)


if __name__ == "__main__":
    main()
