"""
Outlook MCP - OAuth2 Authentication Setup
==========================================
Run this script once to authorize the MCP server to access your Outlook account.
It will open a browser for Microsoft login and store the tokens locally.

Usage:
    python outlook_mcp_auth.py                # Normal mode (opens browser)
    python outlook_mcp_auth.py --no-browser   # Headless mode (no browser)

Environment variables required:
    OUTLOOK_CLIENT_ID      - Azure AD App client ID
    OUTLOOK_CLIENT_SECRET  - Azure AD App client secret
    OUTLOOK_TENANT_ID      - Azure AD tenant ID (or 'common' for multi-tenant)
"""

import os
import sys
import json
import webbrowser
import argparse
from pathlib import Path
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import msal

# Configuration
CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("OUTLOOK_CLIENT_SECRET", "")
TENANT_ID = os.environ.get("OUTLOOK_TENANT_ID", "common")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "http://localhost:5000/callback"
TOKEN_CACHE_PATH = Path.home() / ".outlook_mcp_token_cache.json"

SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Calendars.Read",
    "https://graph.microsoft.com/Calendars.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]


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

    if not CLIENT_ID or not CLIENT_SECRET:
        print("=" * 60)
        print("ERROR: Environment variables not set!")
        print("=" * 60)
        print()
        print("Please set the following environment variables:")
        print()
        print("  export OUTLOOK_CLIENT_ID='your-client-id'")
        print("  export OUTLOOK_CLIENT_SECRET='your-client-secret'")
        print("  export OUTLOOK_TENANT_ID='your-tenant-id'  # or 'common'")
        print()
        print("To get these values:")
        print("  1. Go to https://entra.microsoft.com")
        print("  2. Navigate to: Identity > Applications > App registrations")
        print("  3. Click 'New registration'")
        print("  4. Name: 'Outlook MCP Server'")
        print("  5. Supported account types: pick your preference")
        print("  6. Redirect URI: Web → http://localhost:5000/callback")
        print("  7. After creation, copy the Application (client) ID")
        print("  8. Go to 'Certificates & secrets' → New client secret")
        print("  9. Go to 'API permissions' → Add permission → Microsoft Graph:")
        for scope in SCOPES:
            name = scope.split("/")[-1]
            print(f"     - {name} (Delegated)")
        print("  10. Click 'Grant admin consent' (if applicable)")
        print()
        sys.exit(1)

    # Initialize MSAL
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())

    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY,
        token_cache=cache,
    )

    # Check if we already have a valid token
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            print("✅ Already authenticated! Token is still valid.")
            print(f"   Account: {accounts[0].get('username', 'unknown')}")
            print(f"   Token cache: {TOKEN_CACHE_PATH}")
            if cache.has_state_changed:
                TOKEN_CACHE_PATH.write_text(cache.serialize())
            return

    # Start auth code flow
    flow = app.initiate_auth_code_flow(
        scopes=SCOPES,
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
        print("  4. The browser will redirect to localhost:5000/callback")
        print("  5. If on a different device, copy the FULL redirect URL")
        print("     (including http://localhost:5000/callback?code=...)")
        print("  6. The script will wait for the callback here...")
        print()
    else:
        print("Opening browser for Microsoft login...")
        print(f"If browser doesn't open, visit:\n{auth_url}")
        print()
        webbrowser.open(auth_url)

    # Get authorization code
    if args.code:
        # Manual mode: user provides code or full URL
        print()
        print("Using manually provided authorization code/URL...")

        # Check if it's a full URL or just a code
        if args.code.startswith("http"):
            # Full URL provided
            parsed = urlparse(args.code)
            auth_response = {k: v[0] for k, v in parse_qs(parsed.query).items()}
        else:
            # Just the code provided
            auth_response = {"code": args.code}
    else:
        # Callback server mode
        print("Waiting for authorization callback on http://localhost:5000 ...")
        print()
        print("💡 TIP: If authorizing from a different device, you can also run:")
        print(f"   python outlook_mcp_auth.py --code '<paste-full-callback-url>'")
        print()

        server = HTTPServer(("localhost", 5000), CallbackHandler)

        while CallbackHandler.auth_code is None:
            server.handle_request()

        server.server_close()

        # Complete the flow
        parsed = urlparse(f"http://localhost:5000{CallbackHandler.full_url}")
        auth_response = {k: v[0] for k, v in parse_qs(parsed.query).items()}

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
