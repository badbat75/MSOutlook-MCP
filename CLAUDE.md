# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Outlook MCP Server - A Model Context Protocol server that connects Claude to Microsoft Outlook via Microsoft Graph API. Provides full access to email and calendar operations through 19 MCP tools.

**Core Architecture:**
- **`MCPServer` from the official `mcp` SDK 2.x** (`mcp.server.mcpserver`; the 1.x `FastMCP` import no longer exists) for tool registration and server lifecycle
- **MSAL (Microsoft Authentication Library)** for OAuth2 with automatic token refresh
- **Microsoft Graph API v1.0** for all Outlook operations
- **Async/await** throughout using httpx for HTTP client
- **Two transports, one entry point:** stdio (credentials from `OUTLOOK_*` env vars) or streamable HTTP (credentials from `X-Outlook-*` request headers), selected by the external `outlook_mcp.toml` file

## Project Structure

```
OutlookMCP/
├── outlook_mcp/                # Core package
│   ├── __init__.py
│   ├── app.py                  # The MCPServer instance + lifespan (GraphClientPool)
│   ├── server.py               # Entry point: argparse, .env, config, mcp.run()
│   ├── config.py               # outlook_mcp.toml loader: transport, bind_host, bind_port
│   ├── env.py                  # .env loader + PROJECT_ROOT, shared by every entry point
│   ├── credentials.py          # Credentials, env/header readers, GraphClientPool, get_graph()
│   ├── auth.py                 # AuthManager + GraphClient + shared MSAL token cache
│   ├── folders.py              # Well-known aliases, name lookup, folder tree rendering
│   ├── attachments.py          # Inline (<=3MB) and upload-session attachment writing
│   ├── helpers.py              # Formatting, error handling, $filter validation
│   ├── models.py               # Pydantic input models
│   └── tools/                  # The 19 @mcp.tool() definitions
│       ├── __init__.py         # Imports the three modules = registers every tool
│       ├── mail.py             # 11 email tools
│       ├── calendar.py         # 7 calendar tools
│       └── profile.py          # 1 profile tool
├── scripts/
│   ├── setup-env.ps1           # Load .env + activate venv into YOUR shell (Windows)
│   ├── setup-env.sh            # Load .env + activate venv into YOUR shell (macOS/Linux)
│   ├── generate-claude-config.ps1  # Generate Claude Desktop config (Windows)
│   └── generate-claude-config.sh   # Generate Claude Desktop config (macOS/Linux)
├── tests/
│   ├── unit/                   # pytest, no network: env, config, credentials
│   └── integration/            # Hand-run scripts that call the real Graph API
│       ├── test_mcp_server.py  # JSON-RPC over stdio
│       └── test_http_server.py # Streamable HTTP, credentials as headers
├── docs/
│   ├── SETUP.md                # The single setup guide (was QUICKSTART.md)
│   └── SETUP_PERSONAL_ACCOUNTS.md  # Personal account specifics + AADSTS error table
├── contrib/                    # Deployment helpers, not part of the server
│   ├── mcp_call.py             # Minimal stdio client for calling one tool by hand
│   └── openclaw.json           # MCP host config for the openclaw Linux deployment
├── outlook_mcp_server.py       # Entry point wrapper (identical to the outlook-mcp command)
├── outlook_mcp_auth.py         # OAuth2 initial authorization (standalone)
├── outlook_mcp.toml.example    # Server configuration template (copy to outlook_mcp.toml, gitignored)
├── pyproject.toml              # Package metadata, dependencies, pytest config
├── requirements.txt            # `-e .` only; versions live in pyproject.toml
├── .env.example                # Credentials template
└── claude_desktop_config_example.json
```

**Where things go.** A new tool goes in `tools/`, never in `server.py`:
server.py is the entry point and nothing else. Anything a tool needs that is
not formatting belongs in its own module (`folders.py`, `attachments.py`),
because the one-file version of this package reached 1441 lines and hid its
seams.

## Authentication Flow

The project uses a **two-script approach** for OAuth2:

1. **Initial Setup** (`outlook_mcp_auth.py`):
   - Run once to authorize the app
   - Supports three authorization modes:
     - **Normal mode** (default): Opens browser, waits for callback on port 5000, allows Ctrl+C to paste URL manually
     - **Headless mode** (`--no-browser`): Skips browser/callback, prompts immediately for manual URL input
     - **Direct mode** (`--code <url>`): Accepts pre-obtained authorization code or callback URL
   - Saves tokens to `~/.outlook_mcp_token_cache.json`

   **Usage examples:**
   ```bash
   # Normal mode - Opens browser automatically
   python outlook_mcp_auth.py

   # Headless mode - For remote/SSH systems
   python outlook_mcp_auth.py --no-browser

   # Direct mode - Provide auth code directly
   python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'
   ```

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

### Environment Variables

**Required in stdio mode (and by `outlook_mcp_auth.py`):**
```bash
OUTLOOK_CLIENT_ID      # Azure AD App client ID
OUTLOOK_CLIENT_SECRET  # Azure AD App client secret
OUTLOOK_TENANT_ID      # Azure AD tenant ID or "common"
```

**Optional:**
```bash
OUTLOOK_DOWNLOAD_PATH  # Custom path for email attachments (default: ~/Downloads/outlook_attachments)
OUTLOOK_MCP_CONFIG     # Path to the TOML server configuration (default: <project root>/outlook_mcp.toml)
OUTLOOK_ENV_FILE       # Path to the .env to read (default: <project root>/.env)
```

**The server reads the `.env` itself** through `outlook_mcp/env.py`, from the
project root, whatever the CWD is. That module is the only `.env` parser in the
codebase: the server, `outlook_mcp_auth.py`, the integration scripts and
`contrib/mcp_call.py` all call `load_project_env()`. Do not add a second one.
Variables already present in the environment always win over the file, so a
Claude Desktop `env` block or a systemd `Environment=` line still overrides it.
`--env-file` / `$OUTLOOK_ENV_FILE` move the location; an explicit path that
does not exist is a hard error (exit code 2), mirroring `--config`.

**In HTTP mode the three credential variables are ignored.** Every request must
carry `X-Outlook-Client-Id`, `X-Outlook-Client-Secret` and (optionally,
default `common`) `X-Outlook-Tenant-Id`; a call without them returns an error
naming the missing headers. `OUTLOOK_DOWNLOAD_PATH` stays server-side in both
modes: a remote caller must never choose where the server writes files.

### Server Configuration File (`outlook_mcp.toml`)

Transport and HTTP bind address are never command-line flags; they live in a
TOML file resolved in this order: `--config PATH`, `$OUTLOOK_MCP_CONFIG`,
`<project root>/outlook_mcp.toml` (CWD-independent). No file means stdio.
An explicit path that does not exist, invalid TOML, or invalid values are hard
errors (exit code 2), never a silent fallback to stdio.

```toml
[server]
transport = "http"        # "stdio" (default) or "http"
bind_host = "0.0.0.0"     # HTTP only
bind_port = 8000          # HTTP only
```

The HTTP endpoint is `http://<bind_host>:<bind_port>/mcp` (streamable HTTP).
Template: `outlook_mcp.toml.example`. The real file is gitignored.

**Setting the credentials:**

```bash
cp .env.example .env    # then fill in the three OUTLOOK_* values
```

That is the whole procedure: the server, the auth script and the tests all read
that file. The `scripts/setup-env.*` helpers are for a different job, getting
the same variables plus the activated venv into an **interactive shell**:

```powershell
. .\scripts\setup-env.ps1       # Windows, dot-sourced
```
```bash
source ./scripts/setup-env.sh   # macOS/Linux
```

They keep their own shell parser for that reason; nothing that runs the server
depends on them.

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
python outlook_mcp_auth.py                # Normal mode (opens browser)
python outlook_mcp_auth.py --no-browser   # Headless mode (for remote systems)

# Run with the transport from outlook_mcp.toml (stdio when the file is absent)
python outlook_mcp_server.py

# Run with an explicit configuration file (e.g. transport = "http")
python outlook_mcp_server.py --config /etc/outlook_mcp/outlook_mcp.toml

# Same program, installed console script
outlook-mcp
```

`outlook_mcp_server.py` and `outlook-mcp` both call `outlook_mcp.server.main()`
and behave identically, including the `.env` load. Keep it that way: putting
setup in the wrapper is what made the two diverge before.

### Claude Desktop Integration
Add to `claude_desktop_config.json` (use the venv Python interpreter):

**Windows:**
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

**macOS/Linux:**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "/path/to/OutlookMCP/venv/bin/python",
      "args": ["/path/to/OutlookMCP/outlook_mcp_server.py"]
    }
  }
}
```

No `env` block is needed: the server reads the project `.env`. Add one only to
override it (a different app registration for this host, for instance), since
environment variables take precedence over the file.

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
# stdio (Windows)
claude mcp add outlook -- C:\path\to\OutlookMCP\venv\Scripts\python.exe outlook_mcp_server.py

# stdio (macOS/Linux)
claude mcp add outlook -- /path/to/OutlookMCP/venv/bin/python outlook_mcp_server.py

# HTTP (remote server running with transport = "http"; headers are sent on every request)
claude mcp add --transport http outlook http://server:8000/mcp \
  --header "X-Outlook-Client-Id: ..." \
  --header "X-Outlook-Client-Secret: ..." \
  --header "X-Outlook-Tenant-Id: ..."
```

## Key Implementation Details

### Module Responsibilities

| Module | Purpose |
|--------|---------|
| `outlook_mcp/app.py` | The `mcp = MCPServer(...)` instance, `app_lifespan` (builds the `GraphClientPool`), `get_config()` / `set_config()`. Tool modules import `mcp` from here, which is what keeps server.py free to import the tools |
| `outlook_mcp/server.py` | Entry point only: `_parse_args()`, `main()`, and the `from . import tools` whose side effect registers them |
| `outlook_mcp/env.py` | `PROJECT_ROOT`, `parse_env_file()`, `resolve_env_path()`, `load_project_env()`, `EnvFileError`. The single `.env` parser |
| `outlook_mcp/config.py` | `ServerConfig` + `load_config()`: TOML file lookup and validation for transport / bind_host / bind_port |
| `outlook_mcp/credentials.py` | `Credentials`, `credentials_from_env()`, `credentials_from_headers()`, `GraphClientPool`, `get_graph()` |
| `outlook_mcp/auth.py` | `AuthManager` (MSAL token lifecycle, accepts a shared token cache), `GraphClient` (async HTTP), `load_token_cache()`, `CredentialsError` |
| `outlook_mcp/folders.py` | `WELL_KNOWN_FOLDERS`, `find_folder_id_by_name()`, `resolve_folder()`, `format_folder_tree()` |
| `outlook_mcp/attachments.py` | `read_attachment_meta()`, `attach_small_file()` (<=3MB inline), `attach_large_file()` (upload session), `attach_files()` |
| `outlook_mcp/helpers.py` | Formatting (`format_email_summary()`, `format_event_summary()`), `handle_graph_error()`, `make_recipients()`, `validate_odata_filter()`, `attachment_download_dir()` |
| `outlook_mcp/models.py` | All Pydantic v2 input models with validation |
| `outlook_mcp/tools/` | The `@mcp.tool()` definitions, split mail / calendar / profile |

### Credential Resolution (`get_graph`)

`get_graph(ctx)` in `credentials.py` decides the credential source from the
transport, not from a flag: over HTTP the SDK attaches the Starlette `Request` to
`ctx.request_context.request`, and the `X-Outlook-*` headers of that request
are authoritative; over stdio the request is `None` and the `OUTLOOK_*`
variables read at startup are used. The lifespan holds a `GraphClientPool`
that creates one `AuthManager`/`GraphClient` per distinct credential set on
first use; all of them share one MSAL token cache (entries are keyed by client
id, one `serialize()` persists them all), so several app registrations can be
served by one HTTP process as long as `outlook_mcp_auth.py` was run for each.

### MCP Tool Categories

**Email Tools (11):**
- `outlook_list_mail` - OData filtering, full-text search, pagination ($top, $skip); `folder="*"` searches across the whole mailbox (all folders), and subfolder display names (e.g. "Centri Estivi") resolve automatically
- `outlook_get_mail` - Full message details including body HTML and attachments metadata
- `outlook_list_attachments` - List attachment metadata (name, size, type, ID)
- `outlook_get_attachment` - Download attachment to configured path (default: ~/Downloads/outlook_attachments/) and return file path (all attachments saved to disk - base64 streaming too heavy for MCP; path configurable via OUTLOOK_DOWNLOAD_PATH env var)
- `outlook_send_mail` - HTML body support, CC/BCC, importance levels
- `outlook_create_draft` - Create draft without sending
- `outlook_reply_mail` - Reply or reply-all with comment
- `outlook_move_mail` - Move to folder by name or well-known folder (inbox, archive, deleteditems, etc.)
- `outlook_delete_mail` - Delete a message or draft; soft-delete to Deleted Items by default, or `permanent=True` to delete irrecoverably
- `outlook_update_mail` - Mark read/unread, flag, categorize
- `outlook_list_folders` - Hierarchical folder structure with message counts; recurses into subfolders (nested, indented) by default, toggle with `include_subfolders`

**Calendar Tools (7):**
- `outlook_list_events` - Date range filtering, expands recurring series; with no `calendar_id` it aggregates events across ALL calendars (Calendar, Birthdays, Your Family, etc.), each tagged with its source calendar name
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
4. Returns parsed JSON, or `{"status": "success"}` for empty-body responses (202 Accepted from `sendMail`, 204 No Content from delete/update, or any response with no body)

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

```bash
# 1. Unit tests: no network, no credentials, always runnable. Run these first.
pytest

# 2. If auth logic changed, re-run auth setup
python outlook_mcp_auth.py

# 3. Integration, stdio (real Graph calls; needs a valid token cache)
python tests/integration/test_mcp_server.py
python tests/integration/test_mcp_server.py --verbose  # Full response output
python tests/integration/test_mcp_server.py --quick    # Handshake + profile only

# 4. Integration, HTTP transport (temporary config on a free port, creds sent as
#    headers, server denied both OUTLOOK_* and the project .env)
python tests/integration/test_http_server.py

# 5. Test via Claude Desktop (restart Claude Desktop to reload the server)
```

`pytest` collects `tests/unit` only (see `[tool.pytest.ini_options]`), so a bare
`pytest` never tries to reach Azure. Pure logic belongs there: anything that can
be tested without the network should be, because the integration scripts stop
being runnable the moment an app registration expires.

A quick check that costs nothing and catches most refactoring mistakes: start
the server over stdio and call `tools/list`. It exercises every import and every
decorator without touching Graph.

### Adding New Tools

1. Define the Pydantic input model in `outlook_mcp/models.py` (inherit from `BaseModel`)
2. Add the `@mcp.tool()` decorated async function to the right module in
   `outlook_mcp/tools/` (mail, calendar or profile). Never to `server.py`
3. Use `get_graph(ctx)` from `..credentials` to access the GraphClient
4. Call the Graph endpoint via `graph.get()`, `graph.post()`, etc.
5. Format the response with helpers from `..helpers`
6. Wrap in try/except and use `handle_graph_error(e)` for Graph errors

Example skeleton:
```python
# In outlook_mcp/models.py:
class MyNewToolInput(BaseModel):
    param: str = Field(description="Parameter description")

# In outlook_mcp/tools/mail.py:
from ..app import mcp
from ..credentials import get_graph
from ..helpers import handle_graph_error
from ..models import MyNewToolInput

@mcp.tool(name="outlook_my_new_tool", description="Tool description")
async def my_new_tool(params: MyNewToolInput, ctx: Context = None) -> str:
    graph = get_graph(ctx)
    try:
        result = await graph.get(f"/me/endpoint", params={"key": params.param})
        return json.dumps(result, indent=2)
    except Exception as e:
        return handle_graph_error(e)
```

A tool in a new module needs one line in `outlook_mcp/tools/__init__.py`, or it
is never registered.

### Debugging

- Server logs go to stderr (the SDK's `MCPServer` configures logging); the HTTP startup banner is printed to stderr too, never to stdout
- Token cache issues: delete `~/.outlook_mcp_token_cache.json` and re-auth
- "No valid token available" with a fresh cache and `AADSTS700016` from MSAL means the app registration behind the client id no longer exists in the directory: that is an Azure-side problem, not a code regression
- Graph API errors: check response body in exception (includes error code and message)
- Rate limiting: Graph returns 429 with Retry-After header (not auto-handled currently)
- The server must never depend on its working directory: a stdio server inherits the CWD of whatever host spawned it, which may not even be traversable by the server's user (e.g. a daemon started from another user's 0700 home). With mcp SDK 1.x this used to crash at startup because FastMCP's pydantic-settings probed `./.env`; SDK 2.x reads no `.env` / `MCP_*` at all, and every path this project resolves itself (`outlook_mcp.toml` via `config.DEFAULT_CONFIG_PATH`, the `.env` via `env.DEFAULT_ENV_PATH`) hangs off `env.PROJECT_ROOT`, which is `__file__`-derived. Keep it that way: no relative path, no `Path.cwd()`

## Microsoft Graph API Quirks

- **OData queries** ($filter, $select, $orderBy) have strict syntax - check Graph docs
- **Pagination** uses `@odata.nextLink` (not implemented in tools - uses $top/$skip instead)
- **Recurrence expansion** for calendar events requires `startDateTime` and `endDateTime` query params
- **Meeting creation** sets `isOnlineMeeting: true` to auto-generate Teams link
- **Folder moves** accept either folder ID or well-known name string
- **Attendee types** are: `required`, `optional`, `resource`
