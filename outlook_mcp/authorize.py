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

Credentials required (from the project .env, or already in the environment):
    OUTLOOK_CLIENT_ID      - Azure AD App client ID
    OUTLOOK_CLIENT_SECRET  - Azure AD App client secret
    OUTLOOK_TENANT_ID      - Azure AD tenant ID (or 'common' for multi-tenant)

This is the only place the authorization code flow lives. It talks to MSAL
directly rather than through AuthManager: that class serves an already
authorized account, and its token acquisition now deliberately forces a
refresh, which is the wrong behaviour for a first sign-in.
"""

import argparse
import json
import os
import sys
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import msal

from .auth import (
    GRAPH_SCOPE_URLS,
    REDIRECT_URI,
    TOKEN_CACHE_PATH,
    authority_for,
    load_token_cache,
)
from .env import load_project_env


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
            error = params.get("error", ["unknown"])[0]
            desc = params.get("error_description", [""])[0]
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


def _print_credentials_help(env_file) -> None:
    """Explain how to get the credentials this command needs."""
    print("=" * 60)
    print("ERROR: Azure AD credentials not set!")
    print("=" * 60)
    print()
    if env_file:
        print(f"Read {env_file}, but OUTLOOK_CLIENT_ID / OUTLOOK_CLIENT_SECRET")
        print("are empty there.")
    else:
        print("No .env file found. Copy .env.example to .env and fill it in,")
        print("or export the variables yourself:")
    print()
    print("  OUTLOOK_CLIENT_ID='your-client-id'")
    print("  OUTLOOK_CLIENT_SECRET='your-client-secret'")
    print("  OUTLOOK_TENANT_ID='your-tenant-id'  # or 'common'")
    print()
    print("To get these values:")
    print("  1. Go to https://entra.microsoft.com")
    print("  2. Navigate to: Identity > Applications > App registrations")
    print("  3. Click 'New registration'")
    print("  4. Name: 'Outlook MCP Server'")
    print("  5. Supported account types: pick your preference")
    print(f"  6. Redirect URI: Web → {REDIRECT_URI}")
    print("  7. After creation, copy the Application (client) ID")
    print("  8. Go to 'Certificates & secrets' → New client secret")
    print("  9. Go to 'API permissions' → Add permission → Microsoft Graph:")
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


def main():
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
    args = parser.parse_args()

    # The same .env the server reads, so authorizing needs no separate setup
    # step. Variables already exported in the shell keep precedence.
    env_file = load_project_env()
    client_id = os.environ.get("OUTLOOK_CLIENT_ID", "")
    client_secret = os.environ.get("OUTLOOK_CLIENT_SECRET", "")
    tenant_id = os.environ.get("OUTLOOK_TENANT_ID", "common")

    if not client_id or not client_secret:
        _print_credentials_help(env_file)
        sys.exit(1)

    # Initialize MSAL
    cache = load_token_cache()

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority_for(tenant_id),
        token_cache=cache,
    )

    # Check if we already have a valid token
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPE_URLS, account=accounts[0])
        if result and "access_token" in result:
            print("✅ Already authenticated! Token is still valid.")
            print(f"   Account: {accounts[0].get('username', 'unknown')}")
            print(f"   Token cache: {TOKEN_CACHE_PATH}")
            if cache.has_state_changed:
                TOKEN_CACHE_PATH.write_text(cache.serialize())
            return

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
        print("🔗 MANUAL AUTHORIZATION REQUIRED")
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
            print("❌ No URL provided. Exiting.")
            sys.exit(1)
        auth_response = _auth_response_from(callback_url)

    else:
        # Normal mode: callback server with Ctrl+C fallback
        print("Waiting for authorization callback on http://localhost:5000 ...")
        print()
        print("💡 TIP: If the callback doesn't work, press Ctrl+C to paste manually")
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
                print("❌ No URL provided. Exiting.")
                sys.exit(1)
            auth_response = _auth_response_from(callback_url)

    result = app.acquire_token_by_auth_code_flow(flow, auth_response)

    if "access_token" in result:
        # Save cache
        TOKEN_CACHE_PATH.write_text(cache.serialize())

        print()
        print("✅ Authentication successful!")
        print(f"   Token cache saved to: {TOKEN_CACHE_PATH}")
        print(f"   Scopes granted: {', '.join(result.get('scope', []))}")
        print()
        print("You can now start the MCP server:")
        print("   python outlook_mcp_server.py")
    else:
        print()
        print("❌ Authentication failed!")
        print(f"   Error: {result.get('error', 'unknown')}")
        print(f"   Description: {result.get('error_description', 'N/A')}")
        sys.exit(1)


if __name__ == "__main__":
    main()
