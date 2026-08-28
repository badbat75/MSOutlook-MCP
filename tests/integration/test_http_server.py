r"""
Outlook MCP Server - HTTP transport integration test
=====================================================
Starts the server with a temporary outlook_mcp.toml (transport = "http") and
talks to it over streamable HTTP with plain httpx, sending an X-Auth-Email
header on every request, exactly as the reverse proxy in front of a real
deployment would. No credential is sent: a passing run proves the server serves
one enrolled user's mailbox off its own configured app registration.

Requires the app registration in the project outlook_mcp.toml and a valid token
cache (run outlook_mcp_auth.py once). The test enrolls a throwaway identity by
copying that cache to the per-user path, and removes it again afterwards.

Talks to the real Microsoft Graph. Run it by hand; pytest only collects
tests/unit.

Usage:
    python tests\integration\test_http_server.py
    python tests\integration\test_http_server.py --verbose
"""

import json
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import httpx

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
SERVER_SCRIPT = PROJECT_ROOT / "outlook_mcp_server.py"
VENV_PYTHON = PROJECT_ROOT / "venv" / "Scripts" / "python.exe"
if not VENV_PYTHON.exists():
    VENV_PYTHON = PROJECT_ROOT / "venv" / "bin" / "python"

sys.path.insert(0, str(PROJECT_ROOT))
from outlook_mcp.auth import TOKEN_CACHE_PATH, user_cache_path  # noqa: E402
from outlook_mcp.config import load_config  # noqa: E402

STARTUP_TIMEOUT = 30  # seconds to wait for the port to open
REQUEST_TIMEOUT = 45

USER_HEADER = "X-Auth-Email"
# The identity the proxy would assert. Arbitrary: what makes it work is the
# cache file the test puts at its per-user path, not the address itself.
TEST_USER = "integration-test@localhost"
UNKNOWN_USER = "nobody@localhost"


def get_python():
    return str(VENV_PYTHON) if VENV_PYTHON.exists() else sys.executable


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def wait_for_port(port: int, proc: subprocess.Popen, timeout: float):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if proc.poll() is not None:
            raise RuntimeError(
                f"server exited early (code {proc.returncode}):\n{proc.stderr.read()}"
            )
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return
        except OSError:
            time.sleep(0.2)
    raise RuntimeError(f"server did not open port {port} within {timeout}s")


class HttpMCPClient:
    """Minimal streamable-HTTP JSON-RPC client (handles JSON and SSE replies)."""

    def __init__(self, url: str, headers: dict, verbose: bool = False):
        self.url = url
        self.extra_headers = headers
        self.verbose = verbose
        self.session_id = None
        self._id = 0
        self._http = httpx.Client(timeout=REQUEST_TIMEOUT)

    def close(self):
        if self.session_id:
            try:
                self._http.delete(self.url, headers=self._headers())
            except httpx.HTTPError:
                pass
        self._http.close()

    def _headers(self):
        h = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
            **self.extra_headers,
        }
        if self.session_id:
            h["Mcp-Session-Id"] = self.session_id
        return h

    @staticmethod
    def _parse_body(response: httpx.Response):
        ctype = response.headers.get("content-type", "")
        if ctype.startswith("text/event-stream"):
            # One or more "data: {...}" lines; the last JSON-RPC response wins.
            messages = [
                json.loads(line[5:].strip())
                for line in response.text.splitlines()
                if line.startswith("data:")
            ]
            responses = [m for m in messages if "id" in m]
            return responses[-1] if responses else (messages[-1] if messages else None)
        if response.content:
            return response.json()
        return None

    def send(self, method: str, params=None):
        self._id += 1
        msg = {"jsonrpc": "2.0", "id": self._id, "method": method}
        if params is not None:
            msg["params"] = params
        if self.verbose:
            print(f"  --> {json.dumps(msg)}")
        response = self._http.post(self.url, headers=self._headers(), json=msg)
        if "mcp-session-id" in response.headers:
            self.session_id = response.headers["mcp-session-id"]
        if response.status_code >= 400:
            raise RuntimeError(f"HTTP {response.status_code}: {response.text[:500]}")
        body = self._parse_body(response)
        if self.verbose:
            print(f"  <-- {json.dumps(body)[:500]}")
        return body

    def notify(self, method: str, params=None):
        msg = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            msg["params"] = params
        response = self._http.post(self.url, headers=self._headers(), json=msg)
        if response.status_code >= 400:
            raise RuntimeError(f"HTTP {response.status_code}: {response.text[:500]}")

    def initialize(self):
        result = self.send("initialize", {
            "protocolVersion": "2025-06-18",
            "capabilities": {},
            "clientInfo": {"name": "outlook-http-test", "version": "1.0"},
        })
        assert "result" in result, f"initialize failed: {result}"
        self.notify("notifications/initialized")
        return result["result"]

    def call_tool(self, name: str, arguments: dict) -> str:
        result = self.send("tools/call", {"name": name, "arguments": arguments})
        assert "result" in result, f"tools/call failed: {result}"
        content = result["result"].get("content", [])
        return "\n".join(c.get("text", "") for c in content if c.get("type") == "text")


def enroll_test_user() -> Path:
    """Give TEST_USER the grant the shared cache already holds.

    The per-user cache is an ordinary MSAL cache at a path derived from the
    address, so copying is all "enrolling" means here. Returns the path so the
    caller can remove it again: this test must not leave a mailbox authorized
    under an identity nobody owns.
    """
    if not TOKEN_CACHE_PATH.exists():
        print(f"ERROR: no token cache at {TOKEN_CACHE_PATH}. Run outlook_mcp_auth.py first.")
        sys.exit(1)
    path = user_cache_path(TEST_USER)
    if path.exists():
        print(f"ERROR: {path} already exists; refusing to overwrite it.")
        sys.exit(1)
    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(TOKEN_CACHE_PATH, path)
    return path


def main():
    verbose = "--verbose" in sys.argv or "-v" in sys.argv

    # The app registration the server will run as, from the project's own
    # configuration file: the test copies it into the temporary one so the
    # subprocess needs nothing from this process.
    config = load_config()
    if not config.has_credentials:
        print("ERROR: no [credentials] in outlook_mcp.toml; the server has nothing to run as.")
        sys.exit(1)

    port = free_port()
    url = f"http://127.0.0.1:{port}/mcp"
    identity = {USER_HEADER: TEST_USER}
    user_cache = enroll_test_user()

    with tempfile.TemporaryDirectory() as tmp:
        config_path = Path(tmp) / "outlook_mcp.toml"
        config_path.write_text(
            f'[server]\ntransport = "http"\nbind_host = "127.0.0.1"\nbind_port = {port}\n'
            f'\n[credentials]\n'
            f'client_id = "{config.client_id}"\n'
            f'client_secret = "{config.client_secret}"\n'
            f'tenant_id = "{config.tenant_id}"\n'
            # Turns on the attachment download route, which check_download_route
            # then probes. Also turns on browser enrollment, which nothing here
            # visits.
            f'\n[auth]\npublic_url = "http://127.0.0.1:{port}"\n',
            encoding="utf-8",
        )

        print("=" * 60)
        print("Outlook MCP Server - HTTP Transport Test")
        print("=" * 60)
        print(f"Python:  {get_python()}")
        print(f"Server:  {SERVER_SCRIPT}")
        print(f"URL:     {url}")
        print(f"Config:  {config_path}")
        print(f"User:    {TEST_USER} ({user_cache.name})")
        print()

        proc = subprocess.Popen(
            [get_python(), str(SERVER_SCRIPT), "--config", str(config_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
        )

        passed = failed = 0
        try:
            wait_for_port(port, proc, STARTUP_TIMEOUT)

            checks = [
                ("Initialize + tools/list", check_initialize),
                ("Get Profile (identity from the proxy)", check_profile_as_user),
                ("Missing identity header -> clear error", check_missing_identity),
                ("Unenrolled user -> refused, not app-only", check_unenrolled_user),
                ("Attachment route -> refuses what it did not mint", check_download_route),
                ("List Mail (identity from the proxy)", check_list_mail),
            ]
            for i, (name, fn) in enumerate(checks, 1):
                print(f"  [{i}/{len(checks)}] {name}...", end=" ", flush=True)
                try:
                    detail = fn(url, identity, verbose)
                    print("PASS")
                    if verbose and detail:
                        print(f"        {str(detail)[:300]}")
                    passed += 1
                except Exception as e:  # noqa: BLE001 - report every failure kind
                    print(f"FAIL - {type(e).__name__}: {e}")
                    failed += 1
        finally:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()
            if verbose or failed:
                stderr = proc.stderr.read()
                if stderr:
                    print("\n[server stderr]\n" + stderr[-3000:])
            user_cache.unlink(missing_ok=True)

    print()
    print(f"Results: {passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)


def check_initialize(url, identity, verbose):
    client = HttpMCPClient(url, identity, verbose)
    try:
        info = client.initialize()
        assert info.get("serverInfo", {}).get("name") == "MS_Outlook_MCP", info
        tools = client.send("tools/list")["result"]["tools"]
        names = {t["name"] for t in tools}
        assert "outlook_get_profile" in names, names
        return f"{len(tools)} tools"
    finally:
        client.close()


def check_profile_as_user(url, identity, verbose):
    client = HttpMCPClient(url, identity, verbose)
    try:
        client.initialize()
        text = client.call_tool("outlook_get_profile", {"params": {}})
        assert not text.startswith("Error"), text
        assert "@" in text, f"no email address in profile output: {text[:200]}"
        return text.splitlines()[0]
    finally:
        client.close()


def check_missing_identity(url, identity, verbose):
    client = HttpMCPClient(url, {}, verbose)  # no identity header at all
    try:
        client.initialize()
        text = client.call_tool("outlook_get_profile", {"params": {}})
        assert text.startswith("Error"), f"expected an error, got: {text[:200]}"
        assert USER_HEADER in text, text
        return text[:120]
    finally:
        client.close()


def check_unenrolled_user(url, identity, verbose):
    # The dangerous fallback: with one app registration for everyone, an
    # unenrolled caller must not slide through to client credentials and get a
    # token that acts as the application.
    client = HttpMCPClient(url, {USER_HEADER: UNKNOWN_USER}, verbose)
    try:
        client.initialize()
        text = client.call_tool("outlook_get_profile", {"params": {}})
        assert text.startswith("Error"), f"expected an error, got: {text[:200]}"
        assert UNKNOWN_USER in text, text
        return text[:120]
    finally:
        client.close()


def check_download_route(url, identity, verbose):
    # The route is mounted, runs the same identity policy the tools do, and
    # hands out nothing it did not mint. Minting one for real needs a message
    # with an attachment, which this test does not create.
    base = url.rsplit("/mcp", 1)[0]
    with httpx.Client(timeout=REQUEST_TIMEOUT) as http:
        anonymous = http.get(f"{base}/attachments/not-a-real-token")
        assert anonymous.status_code == 403, (
            f"expected 403 without the identity header, got {anonymous.status_code}"
        )
        unknown = http.get(f"{base}/attachments/not-a-real-token", headers=identity)
        assert unknown.status_code == 404, (
            f"expected 404 for an unknown token, got {unknown.status_code}"
        )
    return "403 without the header, 404 for an unknown token"


def check_list_mail(url, identity, verbose):
    client = HttpMCPClient(url, identity, verbose)
    try:
        client.initialize()
        text = client.call_tool("outlook_list_mail", {"params": {"folder": "inbox", "top": 3}})
        assert not text.startswith("Error"), text
        return text.splitlines()[0]
    finally:
        client.close()


if __name__ == "__main__":
    main()
