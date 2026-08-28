# Setup Guide

Everything needed to go from a fresh clone to a working server. The README
keeps only the short version; this is the file to follow the first time, and
the only place the procedure is written down.

Using a **personal** Microsoft account (outlook.com, hotmail.com, live.com)?
Read [SETUP_PERSONAL_ACCOUNTS.md](SETUP_PERSONAL_ACCOUNTS.md) first: the app
registration has to be created differently, and getting it wrong produces an
error that does not explain itself.

## Prerequisites

- Python 3.11 or newer
- A Microsoft account (personal or work/school)
- An Azure AD app registration (step 1)

---

## 1. Register the app in Azure AD

1. Go to [entra.microsoft.com](https://entra.microsoft.com) and open
   **Identity > Applications > App registrations > New registration**.
2. Name it (for example `Outlook MCP Server`).
3. Supported account types: pick what matches your account. Personal accounts
   need *"Accounts in any organizational directory and personal Microsoft
   accounts"*.
4. Redirect URI: **Web** > `http://localhost:5000/callback`. This exact value
   is what `outlook_mcp_auth.py` listens on.
5. After creation, copy the **Application (client) ID**.
6. **Certificates & secrets > New client secret**: copy the secret *value*
   (not the ID) immediately, it is shown once.
7. **API permissions > Add a permission > Microsoft Graph > Delegated
   permissions**, add all of:
   `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`,
   `Calendars.ReadWrite`, `User.Read`.
8. Grant admin consent if your tenant requires it.

---

## 2. Install

```bash
python -m venv venv

# Windows
venv\Scripts\activate
# macOS/Linux
source venv/bin/activate

pip install -r requirements.txt
```

`requirements.txt` installs the project itself in editable mode, so the
dependency versions come from `pyproject.toml` and the `outlook-mcp` command
lands on your PATH. Add the test dependencies with `pip install -e ".[dev]"`.

---

## 3. Configure the credentials

```bash
# Windows
Copy-Item .env.example .env
# macOS/Linux
cp .env.example .env
```

Fill in the three values from step 1:

```ini
OUTLOOK_CLIENT_ID=your-client-id
OUTLOOK_CLIENT_SECRET=your-client-secret
OUTLOOK_TENANT_ID=common            # or your tenant ID
# OUTLOOK_DOWNLOAD_PATH=C:\Users\YourName\Documents\Outlook_Attachments
```

Both the server and `outlook_mcp_auth.py` read this file themselves, from the
project root, whatever the working directory is. Two rules worth knowing:

- **A variable already set in the environment wins.** The `env` block of a
  Claude Desktop config, or an `Environment=` line in a systemd unit,
  overrides the file rather than the other way round.
- **The location can be moved** with `--env-file PATH` or `$OUTLOOK_ENV_FILE`.
  A path given either way must exist; a typo is an error, not a silent
  fallback to no credentials.

`.env` is gitignored. Never commit it.

---

## 4. Authorize (once)

```bash
python outlook_mcp_auth.py                # opens a browser
python outlook_mcp_auth.py --no-browser   # headless/SSH: prints the URL
python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'
```

Sign in, accept the permissions, and the tokens are written to
`~/.outlook_mcp_token_cache.json`. MSAL refreshes them from then on; re-run
this only if you see `401` errors that persist, or if you add a second app
registration.

If the browser opens but the callback never arrives (a remote machine, a
firewall), press Ctrl+C and paste the full callback URL when prompted.

---

## 5. Choose the transport

The transport and the HTTP listening address live in `outlook_mcp.toml`, never
on the command line, so the server behaves the same whoever starts it. Lookup
order: `--config PATH`, then `$OUTLOOK_MCP_CONFIG`, then `outlook_mcp.toml` in
the project root. **With no file at all the server runs on stdio**, which is
what Claude Desktop and Claude Code want, so most setups can skip this step.

```bash
cp outlook_mcp.toml.example outlook_mcp.toml
```

```toml
[server]
transport = "http"        # "stdio" (default) or "http"
bind_host = "0.0.0.0"     # HTTP only: 127.0.0.1 local, 0.0.0.0 remote
bind_port = 8000          # HTTP only
```

A file that exists but cannot be parsed, or that names an unknown transport or
an out-of-range port, stops the server with exit code 2. It never falls back to
stdio silently: an HTTP deployment that quietly listens on nothing is worse
than one that refuses to start.

The file is gitignored, so each deployment keeps its own. It never holds
credentials.

---

## 6. Start the server

```bash
python outlook_mcp_server.py                       # transport from outlook_mcp.toml
python outlook_mcp_server.py --config /etc/outlook_mcp.toml
outlook-mcp                                        # the installed console script
```

The three forms are the same program.

**Where the credentials come from depends on the transport:**

| Transport | Credentials | Endpoint |
|-----------|-------------|----------|
| `stdio` | `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET`, `OUTLOOK_TENANT_ID` from the environment or the `.env` | process stdin/stdout |
| `http` | `X-Outlook-Client-Id`, `X-Outlook-Client-Secret`, `X-Outlook-Tenant-Id` request headers, sent on every call (the environment is ignored) | `http://<bind_host>:<bind_port>/mcp` |

In HTTP mode `X-Outlook-Tenant-Id` is optional and defaults to `common`; a call
without the two required headers fails with an error naming them. Different
clients may send different app registrations: the server keeps one Graph client
per credential set, all backed by the shared token cache, so run
`outlook_mcp_auth.py` once for every client id you intend to use.

`OUTLOOK_DOWNLOAD_PATH` stays a server-side variable in both modes: a remote
caller must not get to choose where the server writes files.

> Expose an HTTP server only over TLS or behind a reverse proxy. The client
> secret travels in a request header.

---

## 7. Connect a client

### Claude Desktop

Config file location:

- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux:** `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\path\\to\\OutlookMCP\\venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\OutlookMCP\\outlook_mcp_server.py"]
    }
  }
}
```

Use the venv interpreter, not a bare `python`. Credentials come from the
project `.env`; add an `env` block only if you want to override it.

Or generate the file:

```powershell
.\scripts\generate-claude-config.ps1 -Install   # Windows
```
```bash
./scripts/generate-claude-config.sh --install   # macOS/Linux
```

Restart Claude Desktop to reload the server after any change.

### Claude Code

```bash
# stdio, Windows
claude mcp add outlook -- C:\path\to\OutlookMCP\venv\Scripts\python.exe outlook_mcp_server.py

# stdio, macOS/Linux
claude mcp add outlook -- /path/to/OutlookMCP/venv/bin/python outlook_mcp_server.py

# HTTP, remote server
claude mcp add --transport http outlook http://server.example.com:8000/mcp \
  --header "X-Outlook-Client-Id: your-client-id" \
  --header "X-Outlook-Client-Secret: your-client-secret" \
  --header "X-Outlook-Tenant-Id: your-tenant-id"
```

Any MCP client that supports streamable HTTP with custom headers works the same
way; the headers must accompany every request, not only the first.

---

## 8. Verify

```bash
pytest                                        # unit tests, no network
python tests/integration/test_mcp_server.py   # real Graph calls over stdio
python tests/integration/test_http_server.py  # real Graph calls over HTTP
```

`pytest` needs nothing but the dev extra. The two integration scripts need a
valid token cache and a working app registration; see
[README testing notes](../README.md#testing).

---

## Daily usage

The server reads the `.env` on its own, so nothing needs loading before it
starts. The helper scripts exist for the other case: getting the same
variables, plus the activated venv, into **your own shell**.

```powershell
. .\scripts\setup-env.ps1     # Windows
```
```bash
source ./scripts/setup-env.sh # macOS/Linux
```

---

## Troubleshooting

| Symptom | Cause and fix |
|---------|---------------|
| `ModuleNotFoundError` | The venv is not active: `venv\Scripts\activate` |
| `401 Unauthorized` | Token cache stale or missing: re-run `python outlook_mcp_auth.py` |
| `403 Forbidden` | A delegated permission is missing in the app registration (step 1.7) |
| `429` rate limited | Wait the seconds named in the message and retry |
| `Configuration error:` and exit code 2 | `outlook_mcp.toml` or the `--env-file` path is unreadable or invalid; the message names the file |
| Browser callback never arrives | Ctrl+C, then paste the callback URL by hand, or use `--no-browser` |
| `AADSTS...` errors from Microsoft | See the error-by-error table in [SETUP_PERSONAL_ACCOUNTS.md](SETUP_PERSONAL_ACCOUNTS.md#troubleshooting) |
