#!/usr/bin/env python3
"""Lightweight MCP stdio client for MS Outlook MCP server.

Credentials are not handled here: the server reads outlook_mcp.toml itself,
exactly as it does when a real MCP host spawns it.

Usage:
    python contrib/mcp_call.py <tool_name> '<json_arguments>'

Examples:
    python contrib/mcp_call.py outlook_list_mail '{"params":{"folder":"inbox","top":5}}'
    python contrib/mcp_call.py outlook_get_mail '{"params":{"message_id":"AAA..."}}'
    python contrib/mcp_call.py outlook_list_events '{"params":{"top":5}}'
    python contrib/mcp_call.py outlook_get_profile '{"params":{}}'
"""

import json
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SERVER = str(PROJECT_ROOT / "outlook_mcp_server.py")


def venv_python() -> str:
    """The project venv interpreter on either platform, else the current one."""
    for candidate in (
        PROJECT_ROOT / "venv" / "Scripts" / "python.exe",
        PROJECT_ROOT / "venv" / "bin" / "python",
    ):
        if candidate.exists():
            return str(candidate)
    return sys.executable


PYTHON = venv_python()


def call(tool: str, arguments: dict) -> dict:
    proc = subprocess.Popen(
        [PYTHON, SERVER],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )

    def send(msg):
        proc.stdin.write((json.dumps(msg) + "\n").encode())
        proc.stdin.flush()

    def recv():
        line = proc.stdout.readline()
        return json.loads(line) if line else None

    # Handshake
    send({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
        "protocolVersion": "2024-11-05", "capabilities": {},
        "clientInfo": {"name": "binary-mcp-client", "version": "1.0"},
    }})
    recv()  # init response

    send({"jsonrpc": "2.0", "method": "notifications/initialized"})

    # Tool call
    send({"jsonrpc": "2.0", "id": 2, "method": "tools/call",
          "params": {"name": tool, "arguments": arguments}})
    result = recv()

    proc.stdin.close()
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        proc.kill()

    return result


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    # Tool results carry the emoji the server formats its summaries with, and a
    # Windows console is cp1252: printing one raises UnicodeEncodeError and
    # loses the whole response after the call already succeeded. ensure_ascii is
    # off on purpose, to show accented subjects as themselves, so the stream has
    # to tolerate what it cannot render.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")

    tool_name = sys.argv[1]
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {"params": {}}
    result = call(tool_name, args)
    print(json.dumps(result, indent=2, ensure_ascii=False))
