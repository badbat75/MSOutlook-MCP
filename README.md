# Outlook MCP Server

MCP (Model Context Protocol) server that connects Claude to Microsoft Outlook via Microsoft Graph API. Provides full email and calendar management.

## Features

### Email
| Tool | Description |
|------|-------------|
| `outlook_list_mail` | List emails with OData filters, full-text search, pagination |
| `outlook_get_mail` | Full email details (body, headers, attachments) |
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
│   ├── models.py               # Pydantic input models
│   ├── helpers.py              # Formatting and error handling
│   └── server.py               # MCP tool definitions + lifecycle
├── scripts/
│   ├── setup-env.ps1           # Load .env + activate venv (Windows)
│   ├── setup-env.sh            # Load .env + activate venv (macOS/Linux)
│   ├── generate-claude-config.ps1  # Generate Claude Desktop config (Windows)
│   └── generate-claude-config.sh   # Generate Claude Desktop config (macOS/Linux)
├── tests/
│   └── test_mcp_server.py      # Integration tests via JSON-RPC
├── docs/
│   ├── QUICKSTART.md           # Quick start guide
│   └── SETUP_PERSONAL_ACCOUNTS.md  # Personal account setup
├── outlook_mcp_server.py       # Entry point (thin wrapper)
├── outlook_mcp_auth.py         # OAuth2 initial authorization
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
# Windows PowerShell
$env:OUTLOOK_CLIENT_ID = "your-client-id"
$env:OUTLOOK_CLIENT_SECRET = "your-client-secret"
$env:OUTLOOK_TENANT_ID = "your-tenant-id"   # or "common"
```

```bash
# macOS/Linux Bash
export OUTLOOK_CLIENT_ID="your-client-id"
export OUTLOOK_CLIENT_SECRET="your-client-secret"
export OUTLOOK_TENANT_ID="your-tenant-id"   # or "common"
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

### 5. Start the Server

```bash
# For Claude Desktop (stdio)
python outlook_mcp_server.py

# For remote access (HTTP)
python outlook_mcp_server.py --http --port 8000
```

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
        "OUTLOOK_TENANT_ID": "your-tenant-id"
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

```bash
# Windows
claude mcp add outlook -- C:\path\to\OutlookMCP\venv\Scripts\python.exe outlook_mcp_server.py

# macOS/Linux
claude mcp add outlook -- /path/to/OutlookMCP/venv/bin/python outlook_mcp_server.py
```

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
python tests\test_mcp_server.py           # Full test suite
python tests\test_mcp_server.py --quick   # Handshake + profile only
python tests\test_mcp_server.py --verbose # Show full responses
```

**macOS/Linux:**
```bash
source ./scripts/setup-env.sh
python tests/test_mcp_server.py           # Full test suite
python tests/test_mcp_server.py --quick   # Handshake + profile only
python tests/test_mcp_server.py --verbose # Show full responses
```

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
