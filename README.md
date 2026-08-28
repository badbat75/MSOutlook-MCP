# Outlook MCP Server

MCP (Model Context Protocol) server that connects Claude to Microsoft Outlook
via the Microsoft Graph API. 19 tools covering email, calendar and profile,
over stdio or streamable HTTP.

## Features

### Email
| Tool | Description |
|------|-------------|
| `outlook_list_mail` | List emails with OData filters, full-text search, pagination; `folder="*"` searches every folder |
| `outlook_get_mail` | Full email details (body, headers, attachments) |
| `outlook_list_attachments` | List attachment metadata for a message |
| `outlook_get_attachment` | Download attachment to disk (path configurable via env var) |
| `outlook_send_mail` | Send email with HTML, CC/BCC, importance, attachments |
| `outlook_create_draft` | Create draft email without sending |
| `outlook_reply_mail` | Reply or Reply All |
| `outlook_move_mail` | Move email between folders (archive, trash, etc.) |
| `outlook_delete_mail` | Delete a message or draft, to Deleted Items or permanently |
| `outlook_update_mail` | Update read status, categories, flags |
| `outlook_list_folders` | List folders with message counts, nested subfolders included |

### Calendar
| Tool | Description |
|------|-------------|
| `outlook_list_events` | List events in a date range across all calendars (expands recurring series) |
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

## Quick start

```bash
python -m venv venv && venv\Scripts\activate   # macOS/Linux: source venv/bin/activate
pip install -r requirements.txt

cp .env.example .env        # then fill in the Azure AD client id, secret, tenant
python outlook_mcp_auth.py  # sign in once; tokens land in ~/.outlook_mcp_token_cache.json

python outlook_mcp_server.py
```

That covers a local stdio server. The full procedure, including the Azure app
registration and the HTTP transport, is in **[docs/SETUP.md](docs/SETUP.md)**.
Personal Microsoft accounts have their own prerequisites: see
[docs/SETUP_PERSONAL_ACCOUNTS.md](docs/SETUP_PERSONAL_ACCOUNTS.md).

## Transports

One entry point, two transports, chosen by `outlook_mcp.toml` rather than by a
flag. Without that file the server runs on stdio.

| Transport | Credentials | Endpoint |
|-----------|-------------|----------|
| `stdio` | `OUTLOOK_*` variables, from the environment or the project `.env` | process stdin/stdout |
| `http` | `X-Outlook-*` request headers, sent on every call (the environment is ignored) | `http://<bind_host>:<bind_port>/mcp` |

An HTTP deployment carries the client secret in a header: put it behind TLS or
a reverse proxy. See [docs/SETUP.md](docs/SETUP.md#5-choose-the-transport).

---

## Project Structure

```
OutlookMCP/
├── outlook_mcp/                # The package
│   ├── app.py                  # MCPServer instance + lifespan
│   ├── server.py               # Entry point: argparse, .env, config, run
│   ├── config.py               # outlook_mcp.toml loader (transport, bind address)
│   ├── env.py                  # .env loader shared by every entry point
│   ├── credentials.py          # Env/header credential readers + GraphClientPool
│   ├── auth.py                 # AuthManager (MSAL) + GraphClient (httpx)
│   ├── authorize.py            # OAuth2 authorization code flow
│   ├── folders.py              # Mail folder resolution and rendering
│   ├── attachments.py          # Inline and upload-session attachment writing
│   ├── helpers.py              # Formatting, error handling, $filter validation
│   ├── models.py               # Pydantic input models
│   └── tools/                  # The 19 tools
│       ├── mail.py
│       ├── calendar.py
│       └── profile.py
├── tests/
│   ├── unit/                   # pytest, no network
│   └── integration/            # scripts that call the real Graph API
├── scripts/                    # setup-env + Claude Desktop config generators
├── docs/
│   ├── SETUP.md                # Full setup guide
│   └── SETUP_PERSONAL_ACCOUNTS.md
├── contrib/                    # Deployment helpers, not part of the server
├── outlook_mcp_server.py       # Entry point wrapper (same as the outlook-mcp command)
├── outlook_mcp_auth.py         # Wrapper (same as the outlook-mcp-auth command)
├── outlook_mcp.toml.example    # Server configuration template
├── .env.example                # Credentials template
└── pyproject.toml              # Package metadata, dependencies, pytest config
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

```bash
pytest                                          # unit tests, no network, no credentials
python tests/integration/test_mcp_server.py     # stdio, real Graph calls
python tests/integration/test_http_server.py    # HTTP, credentials sent as headers
```

`pytest` collects `tests/unit` only: the .env parser, the TOML config loader
and the credential readers, all pure functions. The two integration scripts
need a valid token cache and a working app registration, so they are run by
hand; `--quick` and `--verbose` are accepted by both.

`test_http_server.py` starts the server on a free port with a temporary
`outlook_mcp.toml`, denies it both the `OUTLOOK_*` variables and the project
`.env`, and forwards your credentials as `X-Outlook-*` headers, so it fails if
anything but the headers is consulted.

---

## Security

- OAuth2 tokens are stored locally in `~/.outlook_mcp_token_cache.json`
- The client secret is never written to logs
- Token refresh is handled automatically by MSAL
- To revoke access: [account.microsoft.com/privacy](https://account.microsoft.com/privacy) > App permissions

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `401 Unauthorized` | Re-run `python outlook_mcp_auth.py` |
| `403 Forbidden` | Check the delegated permissions on the Azure AD app registration |
| `Token expired` | Refresh is automatic; if it persists, re-run auth |
| `Rate limited (429)` | Wait the indicated time and retry |
| `ModuleNotFoundError` | Activate the venv: `venv\Scripts\activate` |
| `Configuration error:` (exit 2) | The named `outlook_mcp.toml` or `.env` path is missing or invalid |
| Browser callback doesn't work | Press Ctrl+C and paste the callback URL manually |
| Remote/SSH system without GUI | Use `python outlook_mcp_auth.py --no-browser` |
| `AADSTS...` from Microsoft | [Error-by-error table](docs/SETUP_PERSONAL_ACCOUNTS.md#troubleshooting) |
