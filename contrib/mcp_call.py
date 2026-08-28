#!/usr/bin/env python3
"""Lightweight MCP stdio client for MS Outlook MCP server.

Usage:
    python mcp_call.py <tool_name> '<json_arguments>'

Examples:
    python mcp_call.py outlook_list_mail '{"params":{"folder":"inbox","top":5}}'
    python mcp_call.py outlook_get_mail '{"params":{"message_id":"AAA..."}}'
    python mcp_call.py outlook_list_events '{"params":{"top":5}}'
    python mcp_call.py outlook_get_profile '{"params":{}}'
"""

import subprocess, json, sys, os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SERVER = os.path.join(SCRIPT_DIR, "outlook_mcp_server.py")
PYTHON = os.path.join(SCRIPT_DIR, "venv", "bin", "python")


def load_dotenv(path: str) -> dict:
    """Minimal .env parser (no external dependency). Returns KEY=VALUE pairs."""
    values = {}
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                values[key.strip()] = val.strip().strip('"').strip("'")
    except FileNotFoundError:
        pass
    return values


# Credentials are read from .env (gitignored), never hardcoded.
_dotenv = load_dotenv(os.path.join(SCRIPT_DIR, ".env"))
ENV = {
    "OUTLOOK_CLIENT_ID": _dotenv.get("OUTLOOK_CLIENT_ID", os.environ.get("OUTLOOK_CLIENT_ID", "")),
    "OUTLOOK_CLIENT_SECRET": _dotenv.get("OUTLOOK_CLIENT_SECRET", os.environ.get("OUTLOOK_CLIENT_SECRET", "")),
    "OUTLOOK_TENANT_ID": _dotenv.get("OUTLOOK_TENANT_ID", os.environ.get("OUTLOOK_TENANT_ID", "common")),
    "PATH": os.environ.get("PATH", "/usr/bin:/bin"),
    "HOME": os.environ.get("HOME", "/home/openclaw"),
}
# Forward optional download path if configured
if _dotenv.get("OUTLOOK_DOWNLOAD_PATH") or os.environ.get("OUTLOOK_DOWNLOAD_PATH"):
    ENV["OUTLOOK_DOWNLOAD_PATH"] = _dotenv.get(
        "OUTLOOK_DOWNLOAD_PATH", os.environ.get("OUTLOOK_DOWNLOAD_PATH", "")
    )


def call(tool: str, arguments: dict) -> dict:
    proc = subprocess.Popen(
        [PYTHON, SERVER],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        env=ENV,
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

    tool_name = sys.argv[1]
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {"params": {}}
    result = call(tool_name, args)
    print(json.dumps(result, indent=2, ensure_ascii=False))
