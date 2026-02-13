r"""
Outlook MCP Server - Automated Integration Test
=================================================
Starts the MCP server as a subprocess and runs JSON-RPC calls via stdio.
Requires valid auth tokens and environment variables.

Usage:
    . .\scripts\setup-env.ps1
    python tests\test_mcp_server.py
    python tests\test_mcp_server.py --verbose       # show full responses
    python tests\test_mcp_server.py --quick         # handshake + profile only
"""

import json
import subprocess
import sys
import os
import time
import threading
import queue
from pathlib import Path
from datetime import datetime, timedelta

TIMEOUT = 45  # seconds per response
PROJECT_ROOT = Path(__file__).parent.parent
SERVER_SCRIPT = PROJECT_ROOT / "outlook_mcp_server.py"
VENV_PYTHON = PROJECT_ROOT / "venv" / "Scripts" / "python.exe"


def get_python():
    """Use venv python if available, else current interpreter."""
    if VENV_PYTHON.exists():
        return str(VENV_PYTHON)
    return sys.executable


class MCPTestClient:
    """Minimal MCP JSON-RPC client over stdio."""

    def __init__(self, verbose=False):
        self.verbose = verbose
        self.process = None
        self._id = 0
        self._lines = queue.Queue()
        self._stderr_lines = queue.Queue()

    def _stdout_worker(self):
        """Background thread that reads stdout lines into a queue."""
        try:
            for line in self.process.stdout:
                self._lines.put(line)
        except ValueError:
            pass  # stream closed

    def _stderr_worker(self):
        """Background thread that reads stderr lines into a queue."""
        try:
            for line in self.process.stderr:
                self._stderr_lines.put(line)
        except ValueError:
            pass  # stream closed

    def start(self):
        python = get_python()
        self.process = subprocess.Popen(
            [python, str(SERVER_SCRIPT)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            bufsize=1,  # line-buffered
            env=os.environ.copy(),
        )
        threading.Thread(target=self._stdout_worker, daemon=True).start()
        threading.Thread(target=self._stderr_worker, daemon=True).start()

    def is_alive(self):
        return self.process and self.process.poll() is None

    def drain_stderr(self):
        """Read any available stderr output (non-blocking)."""
        if not self.process or not self.process.stderr:
            return ""
        lines = []
        try:
            while True:
                line = self._stderr_lines.get_nowait()
                lines.append(line.rstrip())
        except queue.Empty:
            pass
        return "\n".join(lines)

    def stop(self):
        if self.process:
            self.process.stdin.close()
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()

    def _next_id(self):
        self._id += 1
        return self._id

    def send(self, method, params=None):
        """Send a JSON-RPC request and return the parsed response."""
        msg_id = self._next_id()
        msg = {"jsonrpc": "2.0", "id": msg_id, "method": method}
        if params is not None:
            msg["params"] = params

        line = json.dumps(msg) + "\n"
        if self.verbose:
            print(f"  --> {line.strip()}")

        self.process.stdin.write(line)
        self.process.stdin.flush()

        return self._read_response(msg_id)

    def notify(self, method, params=None):
        """Send a JSON-RPC notification (no response expected)."""
        msg = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            msg["params"] = params

        line = json.dumps(msg) + "\n"
        if self.verbose:
            print(f"  --> {line.strip()}")

        self.process.stdin.write(line)
        self.process.stdin.flush()

    def _read_response(self, expected_id):
        """Read lines from queue until we get the response matching our request id."""
        deadline = time.time() + TIMEOUT
        while True:
            remaining = deadline - time.time()
            if remaining <= 0:
                return None

            try:
                line = self._lines.get(timeout=remaining).strip()
            except queue.Empty:
                return None

            if not line:
                continue

            try:
                data = json.loads(line)
            except json.JSONDecodeError:
                continue

            # Skip notifications (no id)
            if "id" not in data:
                if self.verbose:
                    print(f"  <-- (notification) {json.dumps(data)[:200]}")
                continue

            if data.get("id") == expected_id:
                if self.verbose:
                    print(f"  <-- {json.dumps(data)[:500]}")
                return data


# =============================================================================
# Helpers
# =============================================================================

def _extract_text(result):
    """Extract text from MCP tool result (handles both content array and direct result)."""
    if isinstance(result, str):
        return result
    if isinstance(result, dict):
        # FastMCP content array format
        if "content" in result:
            parts = result["content"]
            if isinstance(parts, list):
                return "\n".join(p.get("text", "") for p in parts if p.get("type") == "text")
        # Direct result format
        if "result" in result:
            return result["result"] if isinstance(result["result"], str) else str(result["result"])
    return str(result)


def _assert_tool_success(resp, tool_name):
    """Assert that a tool call returned a successful (non-error) response."""
    assert resp, f"No response from {tool_name}"
    assert "result" in resp, f"No result in {tool_name} response"

    # Check for JSON-RPC level error
    assert "error" not in resp, f"JSON-RPC error from {tool_name}: {resp.get('error')}"

    result = resp["result"]
    text = _extract_text(result)
    assert text, f"Empty response from {tool_name}"

    # Check isError flag from FastMCP
    if isinstance(result, dict) and result.get("isError"):
        raise AssertionError(f"Tool error from {tool_name}: {text[:300]}")

    # Check for error text patterns in the response
    if text.lower().startswith("error:"):
        raise AssertionError(f"Tool returned error from {tool_name}: {text[:300]}")

    return text


# =============================================================================
# Test definitions
# =============================================================================

def test_initialize(client):
    """MCP handshake."""
    resp = client.send("initialize", {
        "protocolVersion": "2024-11-05",
        "capabilities": {},
        "clientInfo": {"name": "test-client", "version": "1.0.0"},
    })
    assert resp and "result" in resp, "No result in initialize response"
    result = resp["result"]
    assert result["protocolVersion"] == "2024-11-05", "Protocol version mismatch"
    assert "tools" in result["capabilities"], "No tools capability"
    assert result["serverInfo"]["name"] == "MS_Outlook_MCP", \
        f"Unexpected server name: {result['serverInfo']['name']}"

    # Send initialized notification
    client.notify("notifications/initialized")
    time.sleep(0.3)
    return result["serverInfo"]


def test_tools_list(client):
    """List all registered tools."""
    resp = client.send("tools/list")
    assert resp and "result" in resp, "No result in tools/list response"
    tools = resp["result"]["tools"]
    assert len(tools) > 0, "No tools registered"

    tool_names = [t["name"] for t in tools]
    expected = [
        "outlook_get_profile",
        "outlook_list_mail",
        "outlook_list_folders",
        "outlook_list_attachments",
        "outlook_get_attachment",
        "outlook_list_events",
        "outlook_list_calendars",
        "outlook_send_mail",
        "outlook_create_draft",
        "outlook_get_mail",
        "outlook_reply_mail",
        "outlook_move_mail",
        "outlook_update_mail",
        "outlook_get_event",
        "outlook_create_event",
        "outlook_update_event",
        "outlook_delete_event",
        "outlook_respond_event",
    ]
    missing = [t for t in expected if t not in tool_names]
    assert not missing, f"Missing tools: {missing}"

    # Verify ctx is NOT in the tool schemas (means injection is working)
    for tool in tools:
        schema_props = tool.get("inputSchema", {}).get("properties", {})
        assert "ctx" not in schema_props, \
            f"Tool '{tool['name']}' still exposes 'ctx' in schema - Context injection not working"

    return tool_names


def test_get_profile(client):
    """Call outlook_get_profile - validates auth is working."""
    resp = client.send("tools/call", {
        "name": "outlook_get_profile",
        "arguments": {},
    })
    return _assert_tool_success(resp, "outlook_get_profile")


def test_list_folders(client):
    """Call outlook_list_folders."""
    resp = client.send("tools/call", {
        "name": "outlook_list_folders",
        "arguments": {"params": {"top": 10}},
    })
    return _assert_tool_success(resp, "outlook_list_folders")


def test_list_mail(client):
    """Call outlook_list_mail with defaults."""
    resp = client.send("tools/call", {
        "name": "outlook_list_mail",
        "arguments": {"params": {"folder": "inbox", "top": 3}},
    })
    return _assert_tool_success(resp, "outlook_list_mail")


def test_list_mail_unread(client):
    """Call outlook_list_mail with isRead filter."""
    resp = client.send("tools/call", {
        "name": "outlook_list_mail",
        "arguments": {"params": {"folder": "inbox", "top": 3, "filter": "isRead eq false"}},
    })
    return _assert_tool_success(resp, "outlook_list_mail (unread)")


def test_list_events(client):
    """Call outlook_list_events for today."""
    today = datetime.now().strftime("%Y-%m-%dT00:00:00")
    tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%dT23:59:59")
    resp = client.send("tools/call", {
        "name": "outlook_list_events",
        "arguments": {"params": {"start_date": today, "end_date": tomorrow}},
    })
    return _assert_tool_success(resp, "outlook_list_events")


def test_list_calendars(client):
    """Call outlook_list_calendars."""
    resp = client.send("tools/call", {
        "name": "outlook_list_calendars",
        "arguments": {"params": {"top": 10}},
    })
    return _assert_tool_success(resp, "outlook_list_calendars")


def test_list_attachments(client):
    """Call outlook_list_attachments on the most recent inbox message."""
    # First get a message ID
    resp = client.send("tools/call", {
        "name": "outlook_list_mail",
        "arguments": {"params": {"folder": "inbox", "top": 1}},
    })
    mail_text = _assert_tool_success(resp, "outlook_list_mail")

    # Extract message ID from response (format: ID: `xxx`)
    import re
    match = re.search(r"ID: `([^`]+)`", mail_text)
    if not match:
        return "SKIP - No message ID found in inbox"

    message_id = match.group(1)

    # Now list attachments
    resp = client.send("tools/call", {
        "name": "outlook_list_attachments",
        "arguments": {"params": {"message_id": message_id}},
    })
    return _assert_tool_success(resp, "outlook_list_attachments")


def test_get_attachment(client):
    """Call outlook_get_attachment on first attachment found."""
    # Get a message with attachments
    resp = client.send("tools/call", {
        "name": "outlook_list_mail",
        "arguments": {"params": {"folder": "inbox", "top": 5, "filter": "hasAttachments eq true"}},
    })
    mail_text = _assert_tool_success(resp, "outlook_list_mail")

    import re
    match = re.search(r"ID: `([^`]+)`", mail_text)
    if not match:
        return "SKIP - No messages with attachments found"

    message_id = match.group(1)

    # List attachments to get attachment ID
    resp = client.send("tools/call", {
        "name": "outlook_list_attachments",
        "arguments": {"params": {"message_id": message_id}},
    })
    att_text = _assert_tool_success(resp, "outlook_list_attachments")

    # Extract first attachment ID
    match = re.search(r"ID: `([^`]+)`", att_text)
    if not match:
        return "SKIP - No attachment ID found"

    attachment_id = match.group(1)

    # Download attachment
    resp = client.send("tools/call", {
        "name": "outlook_get_attachment",
        "arguments": {"params": {"message_id": message_id, "attachment_id": attachment_id}},
    })
    return _assert_tool_success(resp, "outlook_get_attachment")


# =============================================================================
# Runner
# =============================================================================

QUICK_TESTS = [
    ("Initialize", test_initialize),
    ("Tools List", test_tools_list),
    ("Get Profile", test_get_profile),
]

ALL_TESTS = QUICK_TESTS + [
    ("List Folders", test_list_folders),
    ("List Mail", test_list_mail),
    ("List Mail (unread)", test_list_mail_unread),
    ("List Calendars", test_list_calendars),
    ("List Events (today)", test_list_events),
    ("List Attachments", test_list_attachments),
    ("Get Attachment", test_get_attachment),
]


def main():
    verbose = "--verbose" in sys.argv or "-v" in sys.argv
    quick = "--quick" in sys.argv

    tests = QUICK_TESTS if quick else ALL_TESTS

    print("=" * 60)
    print("Outlook MCP Server - Integration Test")
    print("=" * 60)
    print(f"Python:  {get_python()}")
    print(f"Server:  {SERVER_SCRIPT}")
    print(f"Mode:    {'quick' if quick else 'full'} ({len(tests)} tests)")
    print(f"Verbose: {verbose}")
    print()

    # Check env vars
    for var in ("OUTLOOK_CLIENT_ID", "OUTLOOK_CLIENT_SECRET", "OUTLOOK_TENANT_ID"):
        if not os.environ.get(var):
            print(f"ERROR: {var} not set. Run: . .\\scripts\\setup-env.ps1")
            sys.exit(1)

    client = MCPTestClient(verbose=verbose)
    client.start()

    passed = 0
    failed = 0
    errors = []

    try:
        for name, test_fn in tests:
            print(f"  [{passed + failed + 1}/{len(tests)}] {name}...", end=" ", flush=True)
            try:
                result = test_fn(client)
                print("PASS")
                if verbose and result:
                    preview = str(result)[:300].replace("\n", "\n        ")
                    print(f"        {preview}")
                passed += 1
            except AssertionError as e:
                print(f"FAIL - {e}")
                stderr = client.drain_stderr()
                if stderr:
                    print(f"        [stderr] {stderr[:500]}")
                if not client.is_alive():
                    print(f"        [!] Server process died (exit code: {client.process.returncode})")
                errors.append((name, str(e)))
                failed += 1
            except Exception as e:
                print(f"ERROR - {type(e).__name__}: {e}")
                stderr = client.drain_stderr()
                if stderr:
                    print(f"        [stderr] {stderr[:500]}")
                if not client.is_alive():
                    print(f"        [!] Server process died (exit code: {client.process.returncode})")
                errors.append((name, f"{type(e).__name__}: {e}"))
                failed += 1
    finally:
        client.stop()

    # Summary
    print()
    print("=" * 60)
    print(f"Results: {passed} passed, {failed} failed, {passed + failed} total")
    print("=" * 60)

    if errors:
        print()
        print("Failures:")
        for name, err in errors:
            print(f"  - {name}: {err}")

    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
