r"""
Outlook MCP Server - HTTP transport integration test
=====================================================
Starts the server with a temporary outlook_mcp.toml (transport = "http") and
talks to it over streamable HTTP with plain httpx, sending the Azure AD
credentials as X-Outlook-* headers on every request. The server subprocess
gets NO OUTLOOK_* variables, so a passing run proves the headers are the only
credential source in HTTP mode.

Requires a valid token cache (run outlook_mcp_auth.py once) and the
OUTLOOK_CLIENT_ID / OUTLOOK_CLIENT_SECRET / OUTLOOK_TENANT_ID variables in the
*test* environment, which are forwarded as headers.

Usage:
    . .\scripts\setup-env.ps1
    python tests\test_http_server.py
    python tests\test_http_server.py --verbose
"""

import json
import os
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import httpx

PROJECT_ROOT = Path(__file__).parent.parent
SERVER_SCRIPT = PROJECT_ROOT / "outlook_mcp_server.py"
VENV_PYTHON = PROJECT_ROOT / "venv" / "Scripts" / "python.exe"
if not VENV_PYTHON.exists():
    VENV_PYTHON = PROJECT_ROOT / "venv" / "bin" / "python"

STARTUP_TIMEOUT = 30  # seconds to wait for the port to open
REQUEST_TIMEOUT = 45

HEADER_CLIENT_ID = "X-Outlook-Client-Id"
HEADER_CLIENT_SECRET = "X-Outlook-Client-Secret"
HEADER_TENANT_ID = "X-Outlook-Tenant-Id"


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


def main():
    verbose = "--verbose" in sys.argv or "-v" in sys.argv

    creds = {
        HEADER_CLIENT_ID: os.environ.get("OUTLOOK_CLIENT_ID", ""),
        HEADER_CLIENT_SECRET: os.environ.get("OUTLOOK_CLIENT_SECRET", ""),
        HEADER_TENANT_ID: os.environ.get("OUTLOOK_TENANT_ID", "common"),
    }
    if not creds[HEADER_CLIENT_ID] or not creds[HEADER_CLIENT_SECRET]:
        print("ERROR: OUTLOOK_CLIENT_ID / OUTLOOK_CLIENT_SECRET not set. Run: . .\\scripts\\setup-env.ps1")
        sys.exit(1)

    port = free_port()
    url = f"http://127.0.0.1:{port}/mcp"

    # Server environment: everything except the Outlook credentials, so that
    # the only way the tools can work is through the request headers.
    server_env = {k: v for k, v in os.environ.items() if not k.startswith("OUTLOOK_")}

    with tempfile.TemporaryDirectory() as tmp:
        config_path = Path(tmp) / "outlook_mcp.toml"
        config_path.write_text(
            f'[server]\ntransport = "http"\nbind_host = "127.0.0.1"\nbind_port = {port}\n',
            encoding="utf-8",
        )

        print("=" * 60)
        print("Outlook MCP Server - HTTP Transport Test")
        print("=" * 60)
        print(f"Python:  {get_python()}")
        print(f"Server:  {SERVER_SCRIPT}")
        print(f"URL:     {url}")
        print(f"Config:  {config_path}")
        print()

        proc = subprocess.Popen(
            [get_python(), str(SERVER_SCRIPT), "--config", str(config_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            env=server_env,
        )

        passed = failed = 0
        try:
            wait_for_port(port, proc, STARTUP_TIMEOUT)

            checks = [
                ("Initialize + tools/list", check_initialize),
                ("Get Profile (creds via headers)", check_profile_with_headers),
                ("Missing headers -> clear error", check_missing_headers),
                ("List Mail (creds via headers)", check_list_mail),
            ]
            for i, (name, fn) in enumerate(checks, 1):
                print(f"  [{i}/{len(checks)}] {name}...", end=" ", flush=True)
                try:
                    detail = fn(url, creds, verbose)
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

    print()
    print(f"Results: {passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)


def check_initialize(url, creds, verbose):
    client = HttpMCPClient(url, creds, verbose)
    try:
        info = client.initialize()
        assert info.get("serverInfo", {}).get("name") == "MS_Outlook_MCP", info
        tools = client.send("tools/list")["result"]["tools"]
        names = {t["name"] for t in tools}
        assert "outlook_get_profile" in names, names
        return f"{len(tools)} tools"
    finally:
        client.close()


def check_profile_with_headers(url, creds, verbose):
    client = HttpMCPClient(url, creds, verbose)
    try:
        client.initialize()
        text = client.call_tool("outlook_get_profile", {"params": {}})
        assert not text.startswith("Error"), text
        assert "@" in text, f"no email address in profile output: {text[:200]}"
        return text.splitlines()[0]
    finally:
        client.close()


def check_missing_headers(url, creds, verbose):
    client = HttpMCPClient(url, {}, verbose)  # no X-Outlook-* headers at all
    try:
        client.initialize()
        text = client.call_tool("outlook_get_profile", {"params": {}})
        assert text.startswith("Error"), f"expected an error, got: {text[:200]}"
        assert HEADER_CLIENT_ID in text and HEADER_CLIENT_SECRET in text, text
        return text[:120]
    finally:
        client.close()


def check_list_mail(url, creds, verbose):
    client = HttpMCPClient(url, creds, verbose)
    try:
        client.initialize()
        text = client.call_tool("outlook_list_mail", {"params": {"folder": "inbox", "top": 3}})
        assert not text.startswith("Error"), text
        return text.splitlines()[0]
    finally:
        client.close()


if __name__ == "__main__":
    main()
