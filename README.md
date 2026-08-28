# Outlook MCP Server

MCP (Model Context Protocol) server that connects Claude to Microsoft Outlook via Microsoft Graph API. Provides full email and calendar management.

## Features

### Email
| Tool | Description |
|------|-------------|
| `outlook_list_mail` | List emails with OData filters, full-text search, pagination |
| `outlook_get_mail` | Full email details (body, headers, attachments) |
| `outlook_list_attachments` | List attachment metadata for a message |
| `outlook_get_attachment` | Download attachment to disk (path configurable via env var) |
| `outlook_send_mail` | Send email with HTML, CC/BCC, importance |
| `outlook_create_draft` | Create draft email without sending |
| `outlook_reply_mail` | Reply or Reply All |
| `outlook_move_mail` | Move email between folders (archive, trash, etc.) |
| `outlook_update_mail` | Update read status, categories, flags |
| `outlook_list_folders` | List all folders with message counts |

### Calendar
| Tool | Description |
|------|-------------|
| `outlook_list_events` | List events in a date range (expands recurring series) |
| `outlook_get_event` | Full event details with attendees and Teams meeting links |
| `outlook_create_event` | Create event with location, attendees, Teams meeting |
| `outlook_update_event` | Modify or cancel event |
| `outlook_delete_event` | Delete event |
| `outlook_respond_event` | Accept/Decline/Tentative for invitations |
| `outlook_list_calendars` | List all calendars |

### Profile
| Tool | Description |
|------|-------------|
| `outlook_get_profile` | Authenticated user profile info |

---

## Project Structure

```
OutlookMCP/
├── outlook_mcp/                # Core package
│   ├── __init__.py
│   ├── auth.py                 # AuthManager + GraphClient
│   ├── config.py               # outlook_mcp.toml loader (transport, bind host/port)
│   ├── models.py               # Pydantic input models
│   ├── helpers.py              # Formatting and error handling
│   └── server.py               # MCP tool definitions + lifecycle
├── scripts/
│   ├── setup-env.ps1           # Load .env + activate venv (Windows)
│   ├── setup-env.sh            # Load .env + activate venv (macOS/Linux)
│   ├── generate-claude-config.ps1  # Generate Claude Desktop config (Windows)
│   └── generate-claude-config.sh   # Generate Claude Desktop config (macOS/Linux)
├── tests/
│   ├── test_mcp_server.py      # Integration tests via JSON-RPC over stdio
│   └── test_http_server.py     # Integration tests over HTTP (credentials in headers)
├── docs/
│   ├── QUICKSTART.md           # Quick start guide
│   └── SETUP_PERSONAL_ACCOUNTS.md  # Personal account setup
├── outlook_mcp_server.py       # Entry point (thin wrapper)
├── outlook_mcp_auth.py         # OAuth2 initial authorization
├── outlook_mcp.toml.example    # Server configuration template (transport, bind address)
├── requirements.txt
├── pyproject.toml
├── .env.example
├── CLAUDE.md
└── README.md
```

---

## Setup

### 1. Register the App in Azure AD

1. Go to [Microsoft Entra admin center](https://entra.microsoft.com)
2. **Identity > Applications > App registrations > New registration**
3. Configure:
   - **Name:** `Outlook MCP Server`
   - **Supported account types:** choose based on your needs
   - **Redirect URI:** Web > `http://localhost:5000/callback`
4. After creation, copy the **Application (client) ID** and **Directory (tenant) ID**
5. Go to **Certificates & secrets > New client secret** > copy the value
6. Go to **API permissions > Add permission > Microsoft Graph** > Delegated:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.Read`
   - `Calendars.ReadWrite`
   - `User.Read`
7. Click **Grant admin consent** (if you are a tenant admin)

### 2. Install Dependencies

```bash
python -m venv venv
venv\Scripts\activate       # Windows
source venv/bin/activate    # macOS/Linux
pip install -r requirements.txt
```

### 3. Configure Environment Variables

**Windows (PowerShell):**
```powershell
# 1. Create your config from the template
Copy-Item .env.example .env

# 2. Edit .env and fill in your credentials

# 3. Load environment and activate venv
. .\scripts\setup-env.ps1
```

**macOS/Linux (Bash):**
```bash
# 1. Create your config from the template
cp .env.example .env

# 2. Edit .env and fill in your credentials

# 3. Load environment and activate venv
source ./scripts/setup-env.sh
```

**Or set them manually:**

```powershell
# Windows PowerShell - Required variables
$env:OUTLOOK_CLIENT_ID = "your-client-id"
$env:OUTLOOK_CLIENT_SECRET = "your-client-secret"
$env:OUTLOOK_TENANT_ID = "your-tenant-id"   # or "common"

# Optional: Custom download path for attachments
$env:OUTLOOK_DOWNLOAD_PATH = "C:\Users\YourName\Documents\Outlook_Attachments"
```

```bash
# macOS/Linux Bash - Required variables
export OUTLOOK_CLIENT_ID="your-client-id"
export OUTLOOK_CLIENT_SECRET="your-client-secret"
export OUTLOOK_TENANT_ID="your-tenant-id"   # or "common"

# Optional: Custom download path for attachments
export OUTLOOK_DOWNLOAD_PATH="$HOME/Documents/outlook_attachments"
```

### 4. Authorize (First Time)

The authorization script supports three modes:

**Normal mode** (opens browser automatically):
```bash
python outlook_mcp_auth.py
```

**Headless mode** (for remote/SSH systems without GUI):
```bash
python outlook_mcp_auth.py --no-browser
```

**Direct mode** (provide authorization code directly):
```bash
python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'
```

**Normal Mode Workflow:**
- Opens your browser for Microsoft login
- Waits for callback on `http://localhost:5000`
- **TIP:** If callback doesn't work, press Ctrl+C and paste the URL manually

**Headless Mode Workflow:**
- Displays authorization URL to copy
- Paste URL in browser on ANY device (phone, laptop, etc.)
- Copy the callback URL from browser address bar
- Paste it back in the terminal prompt

After authorization, tokens are saved to `~/.outlook_mcp_token_cache.json`.

### 5. Configure the Transport (optional)

The transport and the HTTP listening address are read from `outlook_mcp.toml`
next to `outlook_mcp_server.py` (or from the file named by `--config PATH` /
`$OUTLOOK_MCP_CONFIG`). Without the file the server runs on **stdio**.

```bash
cp outlook_mcp.toml.example outlook_mcp.toml
```

```toml
[server]
transport = "http"        # "stdio" (default) or "http"
bind_host = "0.0.0.0"     # HTTP only: 127.0.0.1 for local clients, 0.0.0.0 for remote
bind_port = 8000          # HTTP only
```

The file is gitignored; each deployment keeps its own copy. It never holds
credentials.

### 6. Start the Server

```bash
python outlook_mcp_server.py                       # transport from outlook_mcp.toml
python outlook_mcp_server.py --config /etc/outlook_mcp.toml
```

**Where credentials come from depends on the transport:**

| Transport | Credentials | Endpoint |
|-----------|-------------|----------|
| `stdio` | `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET`, `OUTLOOK_TENANT_ID` environment variables | process stdin/stdout |
| `http` | `X-Outlook-Client-Id`, `X-Outlook-Client-Secret`, `X-Outlook-Tenant-Id` request headers, sent by the MCP client on every call (environment variables are ignored) | `http://<bind_host>:<bind_port>/mcp` (streamable HTTP) |

In HTTP mode `X-Outlook-Tenant-Id` is optional (defaults to `common`). A call
without the two required headers fails with an error naming them. Different
clients may send different app registrations: the server keeps one Graph
client per credential set, all backed by the shared token cache written by
`outlook_mcp_auth.py`, so run the auth script once for every client id you
intend to use. `OUTLOOK_DOWNLOAD_PATH` stays a server-side environment
variable in both modes (a remote caller must not choose where the server
writes files).

Expose an HTTP server only over TLS or behind a reverse proxy: the client
secret travels in a header.

---

## Claude Desktop Configuration

Add to your Claude Desktop config file (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\path\\to\\OutlookMCP\\venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\OutlookMCP\\outlook_mcp_server.py"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret",
        "OUTLOOK_TENANT_ID": "your-tenant-id",
        "OUTLOOK_DOWNLOAD_PATH": "C:\\Users\\YourName\\Documents\\Outlook_Attachments"
      }
    }
  }
}
```

### Config file location:
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux:** `~/.config/Claude/claude_desktop_config.json`

Or use the config generator:

**Windows:**
```powershell
.\scripts\generate-claude-config.ps1 -Install
```

**macOS/Linux:**
```bash
./scripts/generate-claude-config.sh --install
```

---

## Claude Code Configuration

**stdio (local process, credentials from the environment):**
```bash
# Windows
claude mcp add outlook -- C:\path\to\OutlookMCP\venv\Scripts\python.exe outlook_mcp_server.py

# macOS/Linux
claude mcp add outlook -- /path/to/OutlookMCP/venv/bin/python outlook_mcp_server.py
```

**HTTP (remote server, credentials in headers):**
```bash
claude mcp add --transport http outlook http://server.example.com:8000/mcp \
  --header "X-Outlook-Client-Id: your-client-id" \
  --header "X-Outlook-Client-Secret: your-client-secret" \
  --header "X-Outlook-Tenant-Id: your-tenant-id"
```

Any MCP client that supports streamable HTTP with custom headers works the
same way; the three headers must accompany every request, not only the first.

---

## Usage Examples with Claude

Once configured, you can ask Claude:

- *"Show me my unread emails"*
- *"Send an email to mario@example.com with subject 'Project Proposal'"*
- *"What meetings do I have tomorrow?"*
- *"Create a Teams meeting with the team at 3:00 PM on Monday"*
- *"Archive all newsletter emails"*
- *"Reply to that email saying I confirm"*
- *"Cancel Friday's meeting"*
- *"Accept the meeting invitation for tomorrow"*

---

## Testing

**Windows:**
```powershell
. .\scripts\setup-env.ps1
python tests\test_mcp_server.py           # Full test suite (stdio)
python tests\test_mcp_server.py --quick   # Handshake + profile only
python tests\test_mcp_server.py --verbose # Show full responses
python tests\test_http_server.py          # HTTP transport, credentials sent as headers
```

**macOS/Linux:**
```bash
source ./scripts/setup-env.sh
python tests/test_mcp_server.py           # Full test suite (stdio)
python tests/test_mcp_server.py --quick   # Handshake + profile only
python tests/test_mcp_server.py --verbose # Show full responses
python tests/test_http_server.py          # HTTP transport, credentials sent as headers
```

`test_http_server.py` starts the server on a free port with a temporary
`outlook_mcp.toml`, strips every `OUTLOOK_*` variable from the server's
environment and forwards your credentials as `X-Outlook-*` headers, so it
fails if anything but the headers is consulted.

---

## Security

- OAuth2 tokens are stored locally in `~/.outlook_mcp_token_cache.json`
- Client secret is never exposed in logs
- Token refresh is handled automatically by MSAL
- To revoke access: go to [account.microsoft.com/privacy](https://account.microsoft.com/privacy) > App permissions

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `401 Unauthorized` | Re-run `python outlook_mcp_auth.py` |
| `403 Forbidden` | Check API permissions in Azure AD app registration |
| `Token expired` | Refresh is automatic; if it persists, re-run auth |
| `Rate limited (429)` | Wait the indicated time and retry |
| `ModuleNotFoundError` | Activate venv: `venv\Scripts\activate` |
| Browser callback doesn't work | Press Ctrl+C and paste callback URL manually |
| Remote/SSH system without GUI | Use `python outlook_mcp_auth.py --no-browser` |
