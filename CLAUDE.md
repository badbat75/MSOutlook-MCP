# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Outlook MCP Server - A Model Context Protocol server that connects Claude to Microsoft Outlook via Microsoft Graph API. Provides full access to email and calendar operations through 16 MCP tools.

**Core Architecture:**
- **FastMCP framework** for tool registration and server lifecycle
- **MSAL (Microsoft Authentication Library)** for OAuth2 with automatic token refresh
- **Microsoft Graph API v1.0** for all Outlook operations
- **Async/await** throughout using httpx for HTTP client

## Project Structure

```
OutlookMCP/
├── outlook_mcp/                # Core package
│   ├── __init__.py
│   ├── auth.py                 # AuthManager + GraphClient
│   ├── models.py               # Pydantic input models
│   ├── helpers.py              # Formatting and error handling utilities
│   └── server.py               # MCP tool definitions + lifecycle + entry point
├── scripts/
│   ├── setup-env.ps1           # Load .env + activate venv (Windows)
│   ├── setup-env.sh            # Load .env + activate venv (macOS/Linux)
│   ├── generate-claude-config.ps1  # Generate Claude Desktop config (Windows)
│   └── generate-claude-config.sh   # Generate Claude Desktop config (macOS/Linux)
├── tests/
│   └── test_mcp_server.py      # Integration tests via JSON-RPC over stdio
├── docs/
│   ├── QUICKSTART.md           # Quick start guide
│   └── SETUP_PERSONAL_ACCOUNTS.md  # Personal account setup guide
├── outlook_mcp_server.py       # Entry point (thin wrapper → outlook_mcp.server:main)
├── outlook_mcp_auth.py         # OAuth2 initial authorization (standalone)
├── pyproject.toml              # Package metadata and dependencies
├── requirements.txt            # Pip dependencies
├── .env.example                # Environment variable template
└── claude_desktop_config_example.json
```

## Authentication Flow

The project uses a **two-script approach** for OAuth2:

1. **Initial Setup** (`outlook_mcp_auth.py`):
   - Run once to authorize the app
   - Opens browser for Microsoft login
   - Captures OAuth callback via local HTTP server on port 5000
   - Saves tokens to `~/.outlook_mcp_token_cache.json`

2. **Server Runtime** (`outlook_mcp_server.py` → `outlook_mcp/server.py`):
   - Loads cached tokens on startup via `AuthManager` in `outlook_mcp/auth.py`
   - Handles automatic token refresh via MSAL
   - Falls back to client credentials flow if no cached user token exists
   - Token cache is persisted automatically when state changes

**Critical:** If authentication fails at runtime, the error message will tell users to run `python outlook_mcp_auth.py` again.

## Environment Setup

### Python Virtual Environment

**Always use a virtual environment** to isolate dependencies:

```bash
# Create virtual environment
python -m venv venv

# Activate on Windows
venv\Scripts\activate

# Activate on macOS/Linux
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

**Important:** The virtual environment must be activated before running any scripts or installing dependencies.

### Required Environment Variables
```bash
OUTLOOK_CLIENT_ID      # Azure AD App client ID
OUTLOOK_CLIENT_SECRET  # Azure AD App client secret
OUTLOOK_TENANT_ID      # Azure AD tenant ID or "common"
```

**Setting Environment Variables (recommended approach):**

Use the provided setup script with a `.env` configuration file:

**Windows (PowerShell):**
```powershell
# 1. Copy the example to create your config file
Copy-Item .env.example .env

# 2. Edit .env and fill in your Azure AD credentials

# 3. Run setup script with dot-sourcing to load variables
. .\scripts\setup-env.ps1
```

**macOS/Linux (Bash):**
```bash
# 1. Copy the example to create your config file
cp .env.example .env

# 2. Edit .env and fill in your Azure AD credentials

# 3. Run setup script with source to load variables
source ./scripts/setup-env.sh
```

The setup script will:
- Load environment variables from `.env` file
- Activate the virtual environment automatically
- Validate that all required variables are set
- Display a masked summary of your configuration

**Alternative: Manual environment variable setup:**

```powershell
# Windows PowerShell
$env:OUTLOOK_CLIENT_ID = "your-client-id"
$env:OUTLOOK_CLIENT_SECRET = "your-client-secret"
$env:OUTLOOK_TENANT_ID = "your-tenant-id"

# Windows CMD
set OUTLOOK_CLIENT_ID=your-client-id
set OUTLOOK_CLIENT_SECRET=your-client-secret
set OUTLOOK_TENANT_ID=your-tenant-id

# macOS/Linux bash
export OUTLOOK_CLIENT_ID="your-client-id"
export OUTLOOK_CLIENT_SECRET="your-client-secret"
export OUTLOOK_TENANT_ID="your-tenant-id"
```

### Azure AD App Requirements
The app registration must have these **delegated permissions**:
- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
- `Calendars.Read`, `Calendars.ReadWrite`
- `User.Read`

Redirect URI must be: `http://localhost:5000/callback`

## Running the Server

### Development/Testing
```bash
# Activate virtual environment first!
# Windows: venv\Scripts\activate
# macOS/Linux: source venv/bin/activate

# Initial auth (first time only)
python outlook_mcp_auth.py

# Run in stdio mode (default, for Claude Desktop)
python outlook_mcp_server.py

# Run in HTTP mode (for remote access)
python outlook_mcp_server.py --http --port 8000
```

### Claude Desktop Integration
Add to `claude_desktop_config.json` (use the venv Python interpreter):

**Windows:**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\path\\to\\OutlookMCP\\venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\OutlookMCP\\outlook_mcp_server.py"],
      "env": {
        "OUTLOOK_CLIENT_ID": "...",
        "OUTLOOK_CLIENT_SECRET": "...",
        "OUTLOOK_TENANT_ID": "..."
      }
    }
  }
}
```

**macOS/Linux:**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "/path/to/OutlookMCP/venv/bin/python",
      "args": ["/path/to/OutlookMCP/outlook_mcp_server.py"],
      "env": {
        "OUTLOOK_CLIENT_ID": "...",
        "OUTLOOK_CLIENT_SECRET": "...",
        "OUTLOOK_TENANT_ID": "..."
      }
    }
  }
}
```

Or use the config generator:

**Windows:**
```powershell
.\scripts\generate-claude-config.ps1 -Install
```

**macOS/Linux:**
```bash
./scripts/generate-claude-config.sh --install
```

### Claude Code Integration
```bash
# Windows
claude mcp add outlook -- C:\path\to\OutlookMCP\venv\Scripts\python.exe outlook_mcp_server.py

# macOS/Linux
claude mcp add outlook -- /path/to/OutlookMCP/venv/bin/python outlook_mcp_server.py
```

## Key Implementation Details

### Module Responsibilities

| Module | Purpose |
|--------|---------|
| `outlook_mcp/auth.py` | `AuthManager` (MSAL token lifecycle), `GraphClient` (async HTTP) |
| `outlook_mcp/models.py` | All Pydantic v2 input models with validation |
| `outlook_mcp/helpers.py` | `format_email_summary()`, `format_event_summary()`, `handle_graph_error()`, `make_recipients()` |
| `outlook_mcp/server.py` | FastMCP setup, lifespan context, all 16 `@mcp.tool()` definitions, `main()` entry point |

### MCP Tool Categories

**Email Tools (8):**
- `outlook_list_mail` - OData filtering, full-text search, pagination ($top, $skip)
- `outlook_get_mail` - Full message details including body HTML and attachments metadata
- `outlook_send_mail` - HTML body support, CC/BCC, importance levels
- `outlook_create_draft` - Create draft without sending
- `outlook_reply_mail` - Reply or reply-all with comment
- `outlook_move_mail` - Move to folder by name or well-known folder (inbox, archive, deleteditems, etc.)
- `outlook_update_mail` - Mark read/unread, flag, categorize
- `outlook_list_folders` - Hierarchical folder structure with message counts

**Calendar Tools (7):**
- `outlook_list_events` - Date range filtering, expands recurring series
- `outlook_get_event` - Full event details with attendees and Teams meeting links
- `outlook_create_event` - Supports location, attendees, Teams meeting creation
- `outlook_update_event` - PATCH updates for event modifications
- `outlook_delete_event` - Delete single event
- `outlook_respond_event` - Accept/decline/tentative with optional comment
- `outlook_list_calendars` - List all calendars in account

**Profile Tool (1):**
- `outlook_get_profile` - Current user profile info

### GraphClient Pattern

All Graph API calls go through `GraphClient.request()` in `outlook_mcp/auth.py` which:
1. Gets fresh token via `AuthManager.get_token()` (auto-refreshes if needed)
2. Adds `Authorization: Bearer {token}` header
3. Raises HTTP errors via `httpx.Response.raise_for_status()`
4. Returns JSON or `{"status": "success"}` for 204 responses

Error handling wraps Graph exceptions with `handle_graph_error()` in `outlook_mcp/helpers.py` to provide user-friendly messages.

### Pydantic Models

All tool inputs are defined in `outlook_mcp/models.py` using Pydantic v2 models with:
- Field validation (email addresses, date formats)
- Descriptive field help text for Claude's benefit
- ConfigDict for extra attribute handling
- Custom validators for constrained fields (e.g., importance level, event response type)

### Date/Time Handling

Graph API uses **dateTimeTimeZone** objects:
```json
{
  "dateTime": "2024-01-15T14:00:00",
  "timeZone": "UTC"
}
```

Tools accept ISO 8601 strings and convert to this format via `format_graph_datetime()` in `outlook_mcp/helpers.py`.

### Well-Known Folder Names

Graph API supports aliases like `inbox`, `sentitems`, `deleteditems`, `archive`, `drafts`, `junkemail`. Tools use these directly instead of requiring folder IDs.

## Development Workflow

### Testing Changes

**Windows:**
```powershell
# 0. Load environment
. .\scripts\setup-env.ps1

# 1. Make code changes to files in outlook_mcp/ or outlook_mcp_auth.py

# 2. If auth logic changed, re-run auth setup
python outlook_mcp_auth.py

# 3. Run integration tests
python tests\test_mcp_server.py
python tests\test_mcp_server.py --verbose  # Full response output
python tests\test_mcp_server.py --quick    # Handshake + profile only

# 4. Test via Claude Desktop (restart Claude Desktop to reload server)

# 5. For HTTP mode testing:
python outlook_mcp_server.py --http --port 8000
```

**macOS/Linux:**
```bash
# 0. Load environment
source ./scripts/setup-env.sh

# 1. Make code changes to files in outlook_mcp/ or outlook_mcp_auth.py

# 2. If auth logic changed, re-run auth setup
python outlook_mcp_auth.py

# 3. Run integration tests
python tests/test_mcp_server.py
python tests/test_mcp_server.py --verbose  # Full response output
python tests/test_mcp_server.py --quick    # Handshake + profile only

# 4. Test via Claude Desktop (restart Claude Desktop to reload server)

# 5. For HTTP mode testing:
python outlook_mcp_server.py --http --port 8000
```

### Adding New Tools

1. Define Pydantic input model in `outlook_mcp/models.py` (inherit from `BaseModel`)
2. Add `@mcp.tool()` decorated async function in `outlook_mcp/server.py`
3. Use `_get_graph(ctx)` to access GraphClient
4. Call Graph API endpoint via `graph.get()`, `graph.post()`, etc.
5. Format response using helpers from `outlook_mcp/helpers.py`
6. Wrap in try/except and use `handle_graph_error(e)` for Graph errors

Example skeleton:
```python
# In outlook_mcp/models.py:
class MyNewToolInput(BaseModel):
    param: str = Field(description="Parameter description")

# In outlook_mcp/server.py:
from .models import MyNewToolInput

@mcp.tool(name="outlook_my_new_tool", description="Tool description")
async def my_new_tool(params: MyNewToolInput, ctx: Context = None) -> str:
    graph = _get_graph(ctx)
    try:
        result = await graph.get(f"/me/endpoint", params={"key": params.param})
        return json.dumps(result, indent=2)
    except Exception as e:
        return handle_graph_error(e)
```

### Debugging

- Server logs go to stderr (FastMCP handles logging setup)
- Token cache issues: delete `~/.outlook_mcp_token_cache.json` and re-auth
- Graph API errors: check response body in exception (includes error code and message)
- Rate limiting: Graph returns 429 with Retry-After header (not auto-handled currently)

## Microsoft Graph API Quirks

- **OData queries** ($filter, $select, $orderBy) have strict syntax - check Graph docs
- **Pagination** uses `@odata.nextLink` (not implemented in tools - uses $top/$skip instead)
- **Recurrence expansion** for calendar events requires `startDateTime` and `endDateTime` query params
- **Meeting creation** sets `isOnlineMeeting: true` to auto-generate Teams link
- **Folder moves** accept either folder ID or well-known name string
- **Attendee types** are: `required`, `optional`, `resource`
