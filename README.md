# Outlook MCP Server

MCP (Model Context Protocol) server that connects Claude to Microsoft Outlook
via the Microsoft Graph API. 20 tools covering email, calendar and profile,
over stdio or streamable HTTP.

## Features

### Email
| Tool | Description |
|------|-------------|
| `outlook_list_mail` | List emails with OData filters, full-text search, pagination; `folder="*"` searches every folder |
| `outlook_get_mail` | Full email details (body, headers, attachments) |
| `outlook_list_attachments` | List attachment metadata for a message |
| `outlook_get_attachment` | Download attachment to disk; over HTTP, answers with a one-time link the caller can fetch |
| `outlook_delete_attachment_files` | Remove a message's downloaded files from the server's disk (the mailbox is untouched) |
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

cp outlook_mcp.toml.example outlook_mcp.toml   # then fill in [credentials]
python outlook_mcp_auth.py  # sign in once; tokens land in ~/.outlook_mcp_token_cache.json

python outlook_mcp_server.py
```

That covers a local stdio server. The full procedure, including the Azure app
registration and the HTTP transport, is in **[docs/SETUP.md](docs/SETUP.md)**.
Personal Microsoft accounts have their own prerequisites: see
[docs/SETUP_PERSONAL_ACCOUNTS.md](docs/SETUP_PERSONAL_ACCOUNTS.md).

## Transports

One entry point, two transports, chosen by `outlook_mcp.toml` rather than by a
flag. Without that file the server runs on stdio. The app registration is in
that same file, and **no credential ever travels in a request**: what a call
carries is at most the identity of the user it acts for.

| Transport | Whose mailbox | Token cache | Endpoint |
|-----------|---------------|-------------|----------|
| `stdio` | the account that ran `outlook_mcp_auth.py` | `~/.outlook_mcp_token_cache.json` | process stdin/stdout |
| `http` | the user in `X-Auth-Email`, appended by the reverse proxy | one file per user under `~/.outlook_mcp/caches/` | `http://127.0.0.1:<bind_port>/mcp` |

An HTTP server may only bind loopback, and refuses to start otherwise: the
identity header is trustworthy exactly because the reverse proxy is the only
way in. See [docs/SETUP.md](docs/SETUP.md#multi-user-behind-a-reverse-proxy).

## Deployment

```bash
./scripts/deploy.sh --target root@server.example.com
```

Copies the checkout to a Linux host, builds a virtual environment, installs the
configuration and a systemd unit, and starts the service. `--dir`,
`--service-user`, `--service-name`, `--config` and `--dry-run` are all
accepted; the details are in
[docs/SETUP.md](docs/SETUP.md#deploying-to-a-linux-host). Only what git tracks
is uploaded, so a local `outlook_mcp.toml` never travels by accident.

---

## Project Structure

```
OutlookMCP/
├── outlook_mcp/                # The package
│   ├── app.py                  # MCPServer instance + lifespan
│   ├── server.py               # Entry point: argparse, config, run
│   ├── config.py               # outlook_mcp.toml: the whole configuration
│   ├── credentials.py          # Principal resolution + GraphClientPool
│   ├── auth.py                 # AuthManager (MSAL) + GraphClient (httpx)
│   ├── authorize.py            # OAuth2 authorization code flow
│   ├── enroll.py               # /oauth/login + /oauth/callback (HTTP deployments)
│   ├── downloads.py            # /attachments/<token>: handing a file to a remote caller
│   ├── folders.py              # Mail folder resolution and rendering
│   ├── attachments.py          # Inline and upload-session attachment writing
│   ├── helpers.py              # Formatting, error handling, $filter validation
│   ├── models.py               # Pydantic input models
│   └── tools/                  # The 20 tools
│       ├── mail.py
│       ├── calendar.py
│       └── profile.py
├── tests/
│   ├── unit/                   # pytest, no network
│   └── integration/            # scripts that call the real Graph API
├── scripts/
│   ├── deploy.sh               # Install on a Linux host as a systemd service
│   └── generate-claude-config.*  # Claude Desktop config generators
├── docs/
│   ├── SETUP.md                # Full setup guide
│   └── SETUP_PERSONAL_ACCOUNTS.md
├── contrib/                    # Deployment helpers, not part of the server
│   └── systemd/                # The unit template deploy.sh renders
├── outlook_mcp_server.py       # Entry point wrapper (same as the outlook-mcp command)
├── outlook_mcp_auth.py         # Wrapper (same as the outlook-mcp-auth command)
├── outlook_mcp.toml.example    # The configuration template
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
python tests/integration/test_http_server.py    # HTTP, identity sent as a header
```

`pytest` collects `tests/unit` only: the TOML config loader, the principal
resolution, the token cache layout and the enrollment bookkeeping, all pure
logic. The two integration scripts need a valid token cache and a working app
registration, so they are run by hand; `--quick` and `--verbose` are accepted.

`test_http_server.py` starts the server on a free port with a temporary
`outlook_mcp.toml`, enrols a throwaway identity, and drives it with nothing but
an `X-Auth-Email` header. It checks the refusals too: no header, and a user who
was never enrolled.

---

## Security

- OAuth2 tokens are stored locally in `~/.outlook_mcp_token_cache.json`, or one
  file per user under `~/.outlook_mcp/caches/` in an HTTP deployment, created
  0600 before anything is written into them
- The client secret lives only in `outlook_mcp.toml`, never on the wire, and is
  never written to logs
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
| `Configuration error:` (exit 2) | The named `outlook_mcp.toml` is missing or invalid; the message names the key |
| Refuses to start on a non-loopback bind | Over HTTP the proxy must be the only way in: bind `127.0.0.1` |
| `421 Invalid Host header` behind a proxy | A loopback bind makes the MCP SDK accept only localhost as `Host`. Name the site in `[server].allowed_hosts` |
| `<user> has not authorized this server` | Enrol them: `/oauth/login`, or `outlook-mcp-auth --user <them>` |
| Browser callback doesn't work | Press Ctrl+C and paste the callback URL manually |
| Remote/SSH system without GUI | Use `python outlook_mcp_auth.py --no-browser` |
| `AADSTS...` from Microsoft | [Error-by-error table](docs/SETUP_PERSONAL_ACCOUNTS.md#troubleshooting) |
