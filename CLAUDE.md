# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Outlook MCP Server - A Model Context Protocol server that connects Claude to Microsoft Outlook via Microsoft Graph API. Provides access to email, calendar and contacts through 28 MCP tools.

**Core Architecture:**
- **`MCPServer` from the official `mcp` SDK 2.x** (`mcp.server.mcpserver`; the 1.x `FastMCP` import no longer exists) for tool registration and server lifecycle
- **MSAL (Microsoft Authentication Library)** for OAuth2 with automatic token refresh
- **Microsoft Graph API v1.0** for all Outlook operations
- **Async/await** throughout using httpx for HTTP client
- **Two transports, one entry point:** stdio or streamable HTTP, selected by the external `outlook_mcp.toml` file, which is also where the single Azure AD app registration lives
- **No credential ever travels in a request.** Over stdio the operating system user is the boundary; over HTTP a reverse proxy authenticates the user and appends an identity header, and each user gets their own MSAL token cache

## Project Structure

```
OutlookMCP/
├── outlook_mcp/                # Core package
│   ├── __init__.py
│   ├── app.py                  # The MCPServer instance + lifespan (GraphClientPool)
│   ├── server.py               # Entry point: argparse, config, mcp.run()
│   ├── config.py               # outlook_mcp.toml: PROJECT_ROOT and the whole configuration
│   ├── credentials.py          # Credentials, ProxyAuthPolicy, Principal, GraphClientPool, get_graph()
│   ├── auth.py                 # AuthManager + GraphClient + the token cache layout
│   ├── authorize.py            # The OAuth2 authorization code flow (outlook-mcp-auth)
│   ├── enroll.py               # /oauth/login + /oauth/callback: users enrol themselves
│   ├── downloads.py            # /attachments/<token>, the per-user/per-message file layout
│   ├── folders.py              # Well-known aliases, name lookup, folder tree rendering
│   ├── events.py               # All-day spans, reading an event's times in its own zone
│   ├── contacts.py             # Every contact folder, name matching, birthdays
│   ├── attachments.py          # Inline (<=3MB) and upload-session attachment writing
│   ├── helpers.py              # Formatting, error handling, $filter validation
│   ├── models.py               # Pydantic input models
│   └── tools/                  # The 28 @mcp.tool() definitions
│       ├── __init__.py         # Imports the four modules = registers every tool
│       ├── mail.py             # 12 email tools
│       ├── calendar.py         # 7 calendar tools
│       ├── contacts.py         # 8 contact and contact folder tools
│       └── profile.py          # 1 profile tool
├── scripts/
│   ├── deploy.sh                   # Deploy to a Linux host as a systemd service
│   ├── generate-claude-config.ps1  # Generate Claude Desktop config (Windows)
│   └── generate-claude-config.sh   # Generate Claude Desktop config (macOS/Linux)
├── tests/
│   ├── unit/                   # pytest, no network: config, credentials, auth, enroll, downloads, events, contacts
│   └── integration/            # Hand-run scripts that call the real Graph API
│       ├── test_mcp_server.py  # JSON-RPC over stdio
│       └── test_http_server.py # Streamable HTTP, identity as a header
├── docs/
│   ├── SETUP.md                # The single setup guide (was QUICKSTART.md)
│   └── SETUP_PERSONAL_ACCOUNTS.md  # Personal account specifics + AADSTS error table
├── contrib/                    # Deployment helpers, not part of the server
│   ├── mcp_call.py             # Minimal stdio client for calling one tool by hand
│   ├── openclaw.json           # MCP host config for the openclaw Linux deployment
│   └── systemd/                # The unit template deploy.sh renders and installs
├── outlook_mcp_server.py       # Entry point wrapper (identical to the outlook-mcp command)
├── outlook_mcp_auth.py         # Wrapper (identical to the outlook-mcp-auth command)
├── outlook_mcp.toml.example    # The configuration template (copy to outlook_mcp.toml, gitignored)
├── pyproject.toml              # Package metadata, dependencies, pytest config
└── requirements.txt            # `-e .` only; versions live in pyproject.toml
```

**Where things go.** A new tool goes in `tools/`, never in `server.py`:
server.py is the entry point and nothing else. Anything a tool needs that is
not formatting belongs in its own module (`folders.py`, `attachments.py`),
because the one-file version of this package reached 1441 lines and hid its
seams.

**Nothing in the repo root holds logic.** `outlook_mcp_server.py` and
`outlook_mcp_auth.py` are three-line wrappers over `outlook_mcp.server:main`
and `outlook_mcp.authorize:main`, kept because a daemon, a Claude Desktop
config and every document name those paths. Put code in the package and give it
a console script; a root file that grows a body ends up duplicating what the
package already has, which is how `outlook_mcp_auth.py` came to own a second
copy of the scope list, the authority URL and the token cache path.

## Authentication Flow

The project uses a **two-command approach** for OAuth2:

1. **Initial Setup** (`outlook_mcp_auth.py`, or `outlook-mcp-auth`, both
   `outlook_mcp/authorize.py`):
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

   `--user EMAIL` writes that user's own cache instead of the shared one, for
   an HTTP deployment. It starts from an empty cache on purpose: one file, one
   account, so `get_accounts()` is never ambiguous.

2. **Browser enrollment** (`outlook_mcp/enroll.py`, HTTP deployments only):
   `/oauth/login` and `/oauth/callback`, registered on the HTTP app but
   answering 404 unless `[auth].public_url` is set. The visitor has already
   been authenticated by the reverse proxy, so the routes only collect their
   consent at Microsoft and write their cache. Without this an operator has to
   run `outlook-mcp-auth --user` for every mailbox, which does not scale.

3. **Server Runtime** (`outlook_mcp_server.py` → `outlook_mcp/server.py`):
   - Loads the cache belonging to the caller via `AuthManager` in `outlook_mcp/auth.py`
   - Handles automatic token refresh via MSAL, asking for the scopes of the
     resource each request addresses (see "Scopes per Resource" below)
   - Picks up a cache file a sign-in rewrote while it ran, and never writes its
     own older copy over it
   - Falls back to client credentials only for the single-user (stdio) manager;
     a per-user manager raises instead, so an unenrolled caller is never handed
     a token that acts as the application
   - Token cache is persisted automatically when state changes

**Critical:** If authentication fails at runtime, the error message names the
way out for that caller: `python outlook_mcp_auth.py`, or
`outlook-mcp-auth --user <them>` / `/oauth/login` for a named user.

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

### Configuration (`outlook_mcp.toml`)

**`outlook_mcp/config.py` is the only source of configuration.** There is no
`.env`, no `OUTLOOK_*` variable, and no credential in any request: the one
environment variable left is `$OUTLOOK_MCP_CONFIG`, which says where the file
is and never what is in it. Do not reintroduce an environment fallback: two
sources meant a value could differ between the server and the auth command,
which is how a quoted secret used to work in one process and fail in the next.

Lookup order: `--config PATH`, `$OUTLOOK_MCP_CONFIG`,
`<project root>/outlook_mcp.toml` (CWD-independent). No file at all means a
stdio server with no credentials: it starts, and the first tool call says what
is missing. An explicit path that does not exist, invalid TOML, or invalid
values are hard errors (exit code 2), never a silent fallback.

```toml
[server]
transport = "http"        # "stdio" (default) or "http"
bind_host = "127.0.0.1"   # HTTP only, and loopback only
bind_port = 8000          # HTTP only
allowed_hosts = ["mcp.example.com"]   # HTTP only; the Host a reverse proxy forwards

[credentials]             # the app registration the server runs as, for everyone
client_id = "..."
client_secret = "..."
tenant_id = "common"

[auth]
user_header = "X-Auth-Email"                     # HTTP only
public_url = "https://outlook-mcp.example.com"   # HTTP only; enables /oauth/login and /attachments
cache_dir = "/opt/outlook-mcp/data/caches"       # optional, either transport, absolute

[attachments]
download_path = "~/Downloads/outlook_attachments"   # optional
retention_minutes = 60    # HTTP only; unfetched downloads expire, 0 to keep them
```

The HTTP endpoint is `http://<bind_host>:<bind_port>/mcp` (streamable HTTP).
Template: `outlook_mcp.toml.example`. The real file is gitignored.

`[attachments].download_path` is server side on purpose: a remote caller must
never choose where the server writes files, and never sees a path of its own
choosing served back. `downloads.download_root()` is the one place that reads
it, on every call and never at import time, because the entry point installs
the configuration after the modules are imported.

### The HTTP bind must be loopback (do not relax)

`config._validate_deployment()` refuses to start an HTTP server bound anywhere
but loopback. This is not hardening, it is the premise of the whole mode: over
HTTP a request carries no credential, only the user header, and that header is
believable purely because the reverse proxy replaces whatever the client sent.
A reachable port means anyone can name any mailbox and be served that user's
tokens. If a proxy has to live on another host, tunnel to the loopback port;
do not widen the bind.

The matching duty on the proxy side is `proxy_set_header X-Auth-Email ...` in
**every** location that reaches a mailbox: `/mcp` and `/oauth/`. A location
without it passes the client's own header through. A missing or blank header is
refused, so a forgotten line fails closed.

`/attachments/` is the deliberate exception and must be **unauthenticated**, at
the proxy and in the app alike: see "Downloaded Attachments" below for why a
link the agent cannot follow is not a link, and what carries the weight instead.

### `[server].allowed_hosts`, or: 421 Invalid Host header

A loopback bind makes the MCP SDK turn on DNS-rebinding protection by itself
(`mcp/server/lowlevel/server.py`: `if transport_security is None and host in
("127.0.0.1", "localhost", "::1")`), and it then accepts a Host header of
localhost and nothing else. That rule assumes a proxied server binds a routable
address. This one does not: it binds loopback **and** sits behind a proxy,
because the loopback bind above is the premise of the whole mode. A request
forwarded with the site's own hostname is therefore answered **421 Invalid Host
header**, from the SDK, not from uvicorn or nginx.

`server._transport_security()` is the answer: naming the proxy's hostname in
`[server].allowed_hosts` widens the list while **leaving the protection on**.
Never fix a 421 by disabling the check instead: a loopback server that accepts
any Host is exactly what the protection exists to prevent, and the deployment
would still look fine.

- Each name is registered twice, `host` and `host:*`, because nginx sends the
  bare name with `proxy_set_header Host $host` and name-plus-port with
  `$http_host`; the SDK compares the header literally
- Loopback stays in the list, so a health check on the port itself keeps working
- The key holds Host header values, never URLs. A `https://` prefix would never
  match, so `config.py` refuses one and says what to write instead
- Empty (the default) passes `transport_security=None`, which is identical to
  passing nothing: an unproxied deployment is unaffected

### Azure AD App Requirements
The app registration must have these **delegated permissions**:
- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
- `Calendars.Read`, `Calendars.ReadWrite`
- `User.Read`
- `Contacts.ReadWrite`

For personal accounts consent is dynamic: what a sign-in asks for
(`auth.GRAPH_SCOPES`) is what the user is shown and grants, and on 2026-09-11
the authorize endpoint accepted `Contacts.ReadWrite` for app `576e9b99` before
it was on the registration's list. Keep the list in step all the same: a work
tenant or a `.default` request needs it, and whether it also mattered to an old
grant gaining the scope is open (see "Scopes per Resource").

Redirect URIs: `http://localhost:5000/callback` for `outlook-mcp-auth`, plus
`<public_url>/oauth/callback` when browser enrollment is enabled.

## Running the Server

### Development/Testing
```bash
# Activate virtual environment first!
# Windows: venv\Scripts\activate
# macOS/Linux: source venv/bin/activate

# Initial auth (first time only)
python outlook_mcp_auth.py                # Normal mode (opens browser)
python outlook_mcp_auth.py --no-browser   # Headless mode (for remote systems)
outlook-mcp-auth --user a@b.com           # One user of an HTTP deployment

# Run with the transport from outlook_mcp.toml (stdio when the file is absent)
python outlook_mcp_server.py

# Run with an explicit configuration file (e.g. transport = "http")
python outlook_mcp_server.py --config /etc/outlook_mcp/outlook_mcp.toml

# Same program, installed console script
outlook-mcp
```

`outlook_mcp_server.py` and `outlook-mcp` both call `outlook_mcp.server.main()`
and behave identically, including the configuration load. Keep it that way:
putting setup in the wrapper is what made the two diverge before.

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

No `env` block is needed or useful: the server reads `outlook_mcp.toml` from
the project root. To point one host at a different app registration, add
`"--config", "/path/to/other.toml"` to `args`.

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

# HTTP (remote server behind the reverse proxy that appends the identity header)
claude mcp add --transport http outlook https://outlook-mcp.example.com/mcp
```

The HTTP client sends no credential of its own. It authenticates to the proxy,
and the proxy tells the server whose mailbox to open.

### Deployment (`scripts/deploy.sh` + `contrib/systemd/`)

```bash
./scripts/deploy.sh --target root@host [--dir /opt/outlook-mcp] [--data-dir <dir>/data]
                    [--service-user outlook-mcp] [--service-name outlook-mcp]
                    [--config local.toml] [--port 8000] [--no-restart] [--dry-run]
```

Copies the checkout to a Linux host, builds the venv, installs the config and a
rendered systemd unit, starts the service. The invariants behind it:

- **A unit only makes sense for HTTP.** A stdio server is spawned once per
  client by its MCP host; under systemd it would exit immediately. The script
  refuses a config whose transport is not `http` rather than installing a unit
  that cannot work
- **The config and the data live beside the code, not in it**
  (`<dir>/outlook_mcp.toml`, `<dir>/data/`, `<dir>/app/`). A deploy replaces
  `app/` wholesale so no file survives a rename, which would destroy a config or
  a token cache kept inside it
- **A deployment names both directories it writes to**, `[auth].cache_dir` and
  `[attachments].download_path`, and the deploy refuses a config that puts
  either outside the data dir, since `ReadWritePaths=` grants exactly that one.
  Enrollment then needs nothing but the same `--config` the service runs on.
  The unit still sets `Environment=HOME=<data dir>` and `ProtectHome=no`, but
  only so a stray `~` cannot resolve into a shared service account's home:
  `auth.py` builds its defaults from `Path.home()` at import time, which is
  right on a personal machine and wrong for a deployment
- **Each installation gets its own sign-in.** Never copy a token cache between
  hosts: Entra rotates the refresh token on redemption, so two copies redeeming
  independently invalidate each other, and the failure looks intermittent
- **Only `git ls-files --cached --others --exclude-standard` is uploaded**, so a
  local `outlook_mcp.toml` or `.env` cannot travel by accident. A `--config` is
  staged in a 0700 directory in the SSH user's home, never `/tmp`
- **The config is validated on the host by `load_config()` before the restart**,
  so a bad bind fails the deploy instead of crash-looping. This is also why the
  unit sets `RestartPreventExitStatus=2`: exit 2 is `ConfigError`, and no
  restart will fix it
- The unit template is rendered with `sed` over at-sign delimited names, and the
  deploy warns about any left unsubstituted. Never write such a name in its
  comments: it substitutes there too, and the warning becomes a false alarm

## Key Implementation Details

### Module Responsibilities

| Module | Purpose |
|--------|---------|
| `outlook_mcp/app.py` | The `mcp = MCPServer(...)` instance, `app_lifespan` (builds the `GraphClientPool`, starts and cancels the download reaper), `get_config()` / `set_config()`. Tool modules import `mcp` from here, which is what keeps server.py free to import the tools |
| `outlook_mcp/server.py` | Entry point only: `_parse_args()`, `main()`, and the `from . import tools` whose side effect registers them |
| `outlook_mcp/config.py` | `PROJECT_ROOT`, `ServerConfig` + `load_config()`, `is_loopback()`, `_validate_deployment()`. The whole configuration, and the only place any of it comes from |
| `outlook_mcp/credentials.py` | `Credentials`, `credentials_from_config()`, `ProxyAuthPolicy`, `Principal`, `GraphClientPool`, `current_user()`, `get_graph()`. The pool hands each `AuthManager` a path, never a preloaded cache: see "Scopes per Resource" |
| `outlook_mcp/auth.py` | `AuthManager` (MSAL token lifecycle, one cache and one cache path per principal, `_adopt_rewritten_cache()`), `GraphClient` (async HTTP; `get_bytes()` for media such as a contact's photo), `load_token_cache()` / `save_token_cache()` / `user_cache_path()` / `shared_cache_path()`, `CredentialsError`, `scopes_for()`, and the shared constants `GRAPH_SCOPE_URLS` (what a sign-in asks for) / `CORE_SCOPE_URLS` / `CONTACTS_SCOPE_URLS` / `REDIRECT_URI` / `TOKEN_CACHE_PATH` / `USER_CACHE_DIR` / `authority_for()`. The two path helpers take an optional directory, `None` meaning the home-directory default, so a caller can pass `config.cache_directory` straight through |
| `outlook_mcp/authorize.py` | The OAuth2 authorization code flow: browser, headless, `--code` and `--user` modes. The only place that flow lives |
| `outlook_mcp/enroll.py` | The two enrollment routes and the in-memory table of sign-ins in flight |
| `outlook_mcp/downloads.py` | The `/attachments/<token>` route, the in-memory table of one-time links, `download_root()` / `message_dir()` (where a downloaded file goes), `offer()`, `delete_message_downloads()`, `_consume_file()` (delete on serve) and `sweep()` / `reap_expired_downloads()` (delete on expiry) |
| `outlook_mcp/folders.py` | `WELL_KNOWN_FOLDERS`, `find_folder_id_by_name()`, `resolve_folder()`, `format_folder_tree()` |
| `outlook_mcp/events.py` | `all_day_span()` (what Graph accepts as an all-day event), `describe_all_day()`, `read_event_times()` (an event's times in its own zone, via `Prefer: outlook.timezone`) |
| `outlook_mcp/contacts.py` | `read_all()` (follows `@odata.nextLink`), `ContactFolder` / `contact_folders()` (the tree, default first, with depth and parent), `find_folder()` (an ID, a whole name case and accent blind, or `contacts`), `item_count()` (Graph's `$count`), `read_address_book()` (every folder or the ones given, each contact tagged `folderName`), `folder_name()`, `matches()` (case and accent blind), `month_day()`, `birthday_date()` / `birthday_value()` / `BIRTHDAY_TIME`, `in_deleted_items()`, `copy_of()` / `lost_fields()` / `move_contact()` (copy, check, delete the original; `HalfMoved` when a contact is left twice), `sort_key()` |
| `outlook_mcp/attachments.py` | `read_attachment_meta()`, `attach_small_file()` (<=3MB inline), `attach_large_file()` (upload session), `attach_files()` |
| `outlook_mcp/helpers.py` | Formatting (`format_email_summary()`, `format_event_summary()`, `format_attachment_summary()`, `format_contact_summary()`, `format_contact_details()`), `handle_graph_error()`, `make_recipients()`, `validate_odata_filter()`, `save_attachment_to_disk()` |
| `outlook_mcp/models.py` | All Pydantic v2 input models with validation |
| `outlook_mcp/tools/` | The `@mcp.tool()` definitions, split mail / calendar / contacts / profile |

### Principal Resolution (`get_graph`)

`get_graph(ctx)` in `credentials.py` answers one question, "whose mailbox is
this", from the transport rather than from a flag. The app registration is
always the one in `[credentials]`; what varies is the user:

- **stdio**: `ctx.request_context.request` is `None`. One `Principal` with no
  user, backed by the shared cache: `~/.outlook_mcp_token_cache.json`, or
  `<cache_dir>/shared.json` where one is configured.
- **HTTP**: the SDK attaches the Starlette `Request`, and `ProxyAuthPolicy`
  reads the configured identity header off it. One `Principal` per user, each
  backed by its own `<sha256(address)>.json` under `[auth].cache_dir`, or under
  `~/.outlook_mcp/caches/` when that is unset.

**Per-user caches must stay separate files.** MSAL indexes cache entries by
client id, never by account, so one shared cache under a single app
registration would put every user's account in the same file and
`get_accounts()[0]` would return an arbitrary one: a cross-user token leak. The
isolation is the file, not a lookup key. And do not filter accounts by the
proxy-asserted address instead: the proxy identity and the Microsoft account
are different identity systems and may legitimately differ. What keeps one file
unambiguous is that enrollment always writes a fresh cache.

`AuthManager` carries `user` for the same reason: when it is set, the app-only
client credentials fallback is off. With one app registration serving everyone,
letting an unenrolled caller through to `acquire_token_for_client` would hand
them a token acting as the application itself.

### Secret Verification

**An `AuthManager` never serves a token out of the MSAL cache until AAD has
confirmed its client secret at least once.** MSAL keys cached tokens by client
id and never by secret (`acquire_token_silent` builds its query from
`client_id`, `environment`, `realm`, `home_account_id`).

The bypass this was written for, a caller sending a valid client id with any
string as the secret, is no longer reachable now that credentials come only
from the configuration file. What it still buys is a clear first error:
"wrong secret" and "expired grant" surface as `CredentialsError` on the first
call instead of as a confusing Graph failure later. Cost: one extra token round
trip per principal per process.

How it is enforced in `get_token()`:

- Delegated path: the first call passes `force_refresh=True`, since redeeming
  the refresh token is the request AAD authenticates the secret on. If it
  fails, `CredentialsError` is raised; it must **not** fall through to a cached
  token or to client credentials.
- App-only path: MSAL rejects `force_refresh` on `acquire_token_for_client`, so
  the first call goes through `_unverified_client_app()`, an application with a
  private empty cache that therefore has to reach AAD.

Both paths set `_secret_verified` on success, after which the cache is used
normally. `tests/unit/test_auth.py` pins all of it.

### Scopes per Resource, and Adding a Permission

**A sign-in asks for every scope; a request asks only for the scopes of the
resource it addresses.** `GraphClient.request()` passes `scopes_for(endpoint)`
to `get_token()`: `/me/contacts*` and `/me/contactFolders*` (any casing, an
absolute `@odata.nextLink` included) get `CONTACTS_SCOPE_URLS`, everything
else `CORE_SCOPE_URLS`. Measured on 2026-09-11 against the personal account,
and the reason for all of it:

- A refresh token asked for a single scope its grant lacks fails the **whole**
  request: `invalid_grant`, `AADSTS70000` ("one or more scopes requested are
  unauthorized or expired"); a work account answers `AADSTS65001` with the
  `consent_required` suberror. Requesting the full list on every call would
  have turned mail and calendar off for every grant made before contacts
  existed, and `acquire_token_silent` reports that refusal as `None`, which the
  old code turned into "the client secret is wrong".
- **What happens after the user consents is not settled.** Two observations,
  same account and app, the same day. A refresh token issued before the
  consent, redeemed directly for `Contacts.ReadWrite` two minutes after it,
  still got `AADSTS70000`. Yet the rpi-01 grant, enrolled on 2026-08-28 without
  contacts and never re-enrolled, served the contacts tools forty minutes after
  the consent, with no sign-in there. In between, the user's itadmin agent had
  added the permission to the registration, and on rpi-01 a core-scope refresh
  had just rotated the token. Which of the three (time, the registration, a
  token issued after the consent) made the difference is unknown, and no
  pre-consent token is left to retry. Do not document either outcome as a rule.

So `get_token()` uses `acquire_token_silent_with_error`, and `_lacks_consent()`
turns those codes into a `CredentialsError` that names the missing scope and the
way to grant it (`_consent_message()`). It never falls through to client
credentials, and it does not count as a verified secret. A sign-in on the
installation that reports it always fixes it.

**A sign-in while the server runs must take effect, and must not be undone.**
A sign-in rewrites the cache to grant a new scope, to replace a grant that
stopped working, or to put another account behind the file. `AuthManager`
records the file's `(mtime_ns, size)` when it reads or writes it;
`_adopt_rewritten_cache()` runs at the top of every `get_token()` and reloads
the cache in place when the file changed under it (a deleted file becomes an
empty cache), and `_save_cache()` refuses to write when the file is not the one
it last saw. Before this, a running server kept its first copy until
restarted, and its next refresh wrote that copy back over the sign-in. It also
fixes a user who called a tool before enrolling staying "not authorized" until
a restart. That is why the pool no longer preloads the shared cache at startup:
the copy a manager holds has to be the file as it was when the manager read it.

To add a permission later: add a scope list in `auth.py`, include it in
`GRAPH_SCOPES`, map its resource paths in `_RESOURCE_SCOPE_URLS`, add it to the
registration and the docs. Existing users keep every other tool; the new ones
tell them to sign in again until the scope is theirs.

### Downloaded Attachments (`downloads.py`)

`outlook_get_attachment` writes to the filesystem of the machine the **server**
runs on. Over stdio that is the caller's machine too, so the path in the answer
is the deliverable. Over HTTP it is not, and a path there is useless to whoever
asked: `downloads.py` closes that gap with one route and one file layout.

**The layout is the mechanism, not bookkeeping.** A file lands at
`<download_path>/<sha256(address)>/msg-<sha256(message id)[:16]>/<name>`, and
each level pays for itself:

- The **user** level is the same isolation the token caches have. Without it one
  directory holds everybody's mail attachments, and one person's filenames show
  up in the duplicate suffixes generated for another's. Over stdio it is absent:
  there is one person, and it is their own download directory.
- The **message** level is what makes deletion possible with no index to keep
  in sync. `outlook_delete_attachment_files` is handed a message id, never a
  path, and empties exactly one directory. Recomputing beats recording.

**A download link is an unauthenticated capability that destroys itself.** This
is the one place in the project where a route asks for nothing, and it is a
requirement rather than a concession. The party that has to fetch the file is
the agent that just called the tool; the credential its MCP client presents on
`/mcp` is deliberately one the agent never sees, so a link gated on that
credential can never be followed by the only caller that wants it. Handing the
key to the agent instead would put it in the transcript, which is worse than
anything this route can leak.

So the token is the whole credential, and it is made worthless as fast as
possible:

- `offer()` mints 256 unguessable bits into an in-memory table, naming one file.
- `_take()` pops the ticket, so a second fetch is a 404.
- `_consume_file()` runs as the response's `BackgroundTask`, after the body has
  been streamed off that very file, and **deletes it**. The link in the answer
  is already dead by the time the answer is read.
- `reap_expired_downloads()` sweeps whatever nobody fetched once
  `[attachments].retention_minutes` runs out (60 by default, `0` disables it).
  It walks the filesystem, not the ticket table: a link expires in fifteen
  minutes while the file outlives it, and a restart empties the table entirely.

Do not add a "fetch it again" convenience, and do not make the link reusable:
one fetch plus deletion is what buys the right to skip authentication. And do
not add an identity check back "for safety" without removing the whole scheme:
a check the agent cannot satisfy just breaks downloads, which is the bug this
replaced.

Consequences worth keeping in mind:

- **`/attachments/` must be exempt from the proxy's authentication**, unlike
  `/mcp` and `/oauth/`. A gate there answers 401 to the agent and nothing works.
- The route answers **404 unless `[auth].public_url` is set**, and without it
  the tool falls back to reporting the server-side path and says so. It has no
  other way to know what URL it is reachable on.
- **Retention is HTTP only** (`ServerConfig.retention_seconds` returns `None`
  over stdio). There the file the tool wrote is the answer itself, in the
  caller's own download directory: expiring it would delete the user's file.
- The reaper is started by `app_lifespan` and cancelled with it, because a
  background task has to be cancelled by whatever started it. `app.py` imports
  `downloads` inside the lifespan function: `downloads` registers its route on
  `mcp`, so it imports `app` and cannot be imported by it at module scope.

### MCP Tool Categories

**Email Tools (12):**
- `outlook_list_mail` - OData filtering, full-text search, pagination ($top, $skip); `folder="*"` searches across the whole mailbox (all folders), and subfolder display names (e.g. "Centri Estivi") resolve automatically
- `outlook_get_mail` - Full message details including body HTML and attachments metadata
- `outlook_list_attachments` - List attachment metadata (name, size, type, ID)
- `outlook_get_attachment` - Download attachment to the configured path (default: ~/Downloads/outlook_attachments/), filed per user and per message. Answers with the file path over stdio, and over HTTP with a one-time link at `<public_url>/attachments/<token>` that needs no credential and deletes the file when fetched. Every attachment goes to a file: a base64 data URL of a real attachment is far too heavy to send back through MCP
- `outlook_delete_attachment_files` - Remove what was downloaded from one message (or one named file of it) from the server's filesystem. The mailbox is untouched and the attachment can be fetched again. A convenience over HTTP, where a fetched file is already gone and an unfetched one expires; the only way over stdio
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
- `outlook_create_event` - Supports location, attendees, Teams meeting creation, `is_all_day` and `show_as` (free, tentative, busy, oof, workingElsewhere)
- `outlook_update_event` - PATCH updates for event modifications, including `is_all_day` and `show_as`; see "All-Day Events" below for why it is more than a pass-through
- `outlook_delete_event` - Delete single event
- `outlook_respond_event` - Accept/decline/tentative with optional comment
- `outlook_list_calendars` - List all calendars in account

**Contact Tools (8):**
- `outlook_list_contacts` - Every contact folder, not just the default one; name, email, phone, birthday, folder, ID. `search` matches names and addresses ignoring case and accents; `birthday` takes an ISO date whose year is ignored (or `--MM-DD`), which is how to find the contacts behind a Birthdays calendar entry; `folder` keeps one folder (not the ones inside it). Sorted by name, nameless last; `top`/`skip` page the filtered list
- `outlook_get_contact` - Everything a contact holds, so a duplicate is read before it is deleted
- `outlook_create_contact` - The fields of an update, plus `folder` (a name, an ID, or `contacts`, the default). Needs a name, an address or a number; without `display_name` Outlook composes one from the name parts ("Zz" + "Test MCP Uno" became "Zz Test MCP Uno")
- `outlook_update_contact` - Names, company, email addresses (at most 3), phones (mobile, at most 2 home and 2 business), birthday (`''` removes it), notes. Lists replace; the display name always goes along unless given
- `outlook_delete_contact` - To Deleted Items, named in the answer. No permanent option, on purpose
- `outlook_list_contact_folders` - The folder tree, default first, with each folder's contacts, what Graph counts beyond them, and its ID
- `outlook_move_contacts` - Up to 25 IDs into one folder, one at a time: copy with every writable property, check the copy, original to Deleted Items. New ID and creation date; a contact with a photo is refused; an unexpected error stops the call and lists what was not attempted. See "Moving Contacts, and Contact Folders"
- `outlook_delete_contact_folder` - An empty folder only (no contacts, a `$count` of 0, no subfolders), to Deleted Items; never the default one

**Profile Tool (1):**
- `outlook_get_profile` - Current user profile info

### GraphClient Pattern

All Graph API calls go through `GraphClient.request()` in `outlook_mcp/auth.py` which:
1. Gets fresh token via `AuthManager.get_token()` (auto-refreshes if needed)
2. Adds `Authorization: Bearer {token}` header
3. Raises HTTP errors via `httpx.Response.raise_for_status()`
4. Returns parsed JSON, or `{"status": "success"}` for empty-body responses (202 Accepted from `sendMail`, 204 No Content from delete/update, or any response with no body)

`GraphClient.get_bytes()` is the exception: raw content and its media type (a
contact's photo at `.../photo/$value`), and `None` for a 404, which it does not
log. Both go through `_send()`, which logs every other error status together
with Graph's answer.

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

### All-Day Events (`events.py`)

Graph is strict about all-day events, and none of it shows in the event
resource docs. Probed against a real mailbox:

- An all-day event must run **midnight to midnight**, the end exclusive and at
  least 24 hours after the start, both ends in **the same time zone**. Anything
  else is a 400.
- A PATCH carrying `isAllDay`, **true or false**, is refused unless `start`
  comes with it ("Missing parameters: Event.Start"), even when the event already
  sits on midnight. Forwarding the flag alone therefore never works, which is
  why `outlook_update_event` reads the event's current times when the caller
  gives none and sends them along.
- Those times are read in the caller's `timezone`, else the event's own
  (`originalStartTimeZone`), via `Prefer: outlook.timezone`. Never in the UTC
  Graph answers in by default: midnight in Rome is 22:00 the day before in UTC,
  which would make the event all-day on the wrong day.
- `all_day_span()` widens whatever it gets to the days it touches. An end given
  as a date is the last day (inclusive); an end given as a time is an instant,
  so one at midnight is exclusive. The second rule is what makes reading an
  all-day event back and passing its times in again a no-op.
- Once all-day, the mailbox stores the event as a floating date: a GET reports
  it at 00:00 in whatever zone is asked for, `originalStartTimeZone` included.
- `showAs` accepts any casing at Graph; the model normalizes it anyway so the
  schema can carry an `enum`. A new event without it is `busy`, all-day or not.

### Contacts (`contacts.py`)

Written for a request to merge duplicate contacts: three entries on 7 June in
the Birthdays calendar, which shows only the day and the month. Everything below
was measured against the real mailbox on 2026-09-11.

- **`/me/contacts` is the default folder only.** That mailbox also has
  "HUAWEI P40 Pro (contacts synced by Link to Windows)", 135 contacts inside the
  default one: exactly where a duplicate lives. `contact_folders()` starts from
  `/me/contactFolders/contacts` (the well-known name resolves to the default
  folder, which `/me/contactFolders` does not list) and walks `childFolders`.
  `/me/contacts/{id}` does reach a contact in any folder.
- **No name search at Graph.** `$filter` on a contact reaches only
  `emailAddresses/any(a:a/address eq '...')`, so the whole address book is read
  (342 contacts: two pages) and matched here.
- **Birthdays are instants.** Every one of the 21 in the mailbox sits at
  `T11:59:00Z`, what Outlook writes, and the Birthdays calendar shows its UTC
  date. The tools read the UTC date and write `YYYY-MM-DDT11:59:00Z`. A contact
  with a birthday gets a Birthdays entry `"<displayName> birthday"` within
  seconds, and loses it when the birthday is removed or the contact deleted.
- **A Birthdays entry without the " birthday" suffix is not from a contact.**
  The third 7 June entry, plain "Gabriele", matched no contact in any folder:
  it is a yearly all-day series created on 2017-07-03, while an entry generated
  from a contact is created within a second of the contact. That one is for
  `outlook_delete_event`; merging contacts leaves it alone.
- **An update that leaves out `displayName` regenerates it.** A PATCH of
  `givenName` alone turned "Zz Test MCP (temporaneo)" into "Qq Test MCP", as the
  contact resource doc warns. `outlook_update_contact` reads the current one and
  sends it along. Email addresses already on the contact keep their `name`.
- **Phones synced from a phone put odd things in odd fields**: "Gabriele, Amore
  di Papà" has the endearment in `generation` (the suffix). `get_contact` shows
  every populated field for that reason.
- **No permanent delete.** `permanentDelete` purged a contact without a
  birthday, but twice left one with a birthday in Deleted Items, once answering
  404. An option that says "not recoverable" and is not is worse than none.
- **`DELETE` has answered 404 while carrying itself out**, three times in a
  few minutes, on contacts given a birthday seconds before; not reproduced in
  about a dozen tries since, with the server log capturing every Graph error.
  `in_deleted_items()` looks for the contact there by name and creation time
  (the move changes the ID) before a 404 is reported.

### Moving Contacts, and Contact Folders

Written for the next request on the same mailbox: the phone behind "HUAWEI P40
Pro (contacts synced by Link to Windows)", a child of the default folder, no
longer syncs, and its 115 contacts were to be folded into the default folder
before the folder is deleted. Measured on 2026-09-11 with throwaway contacts:

- **Graph can neither move nor copy a contact.** `POST /me/contacts/{id}/move`
  and `.../copy` answer 400 "Resource not found for the segment 'move'". A move
  is a copy made from a full GET (`copy_of()`: every property but the read-only
  ones and the `@odata` annotations), then a soft delete of the original. The
  copy came back equal to the original in every property. Its ID and
  `createdDateTime` are new, and nothing can set them.
- **A personal mailbox echoes the first three addresses** as
  `primaryEmailAddress` / `secondaryEmailAddress` / `tertiaryEmailAddress`,
  undocumented in v1.0. They matched `emailAddresses` in all 277 contacts; a
  POST accepts them, and `copy_of()` leaves them out all the same.
- **A copy carrying a photo vanishes once its original is deleted**, from every
  folder and from Deleted Items, 10 to 45 seconds after the delete: seen three
  times, with nothing but the photo in common. Copies with the phone's category,
  a birthday or email addresses survived every time, whether the original was
  deleted or kept. A photo copy whose original is kept was not tried, and the
  mechanism is unknown. Hence `outlook_move_contacts` leaves a contact with a
  photo where it is and says so. The phone folder has none; the default folder
  has 4. Uploading a photo works on this account (`PUT .../photo/$value`, 200),
  which is how it was measured, but no tool does it: do not make the move carry
  photos without measuring first what happens to them.
- **The copy's birthday gets its own entry** in the Birthdays calendar, and the
  original's entry goes with the original.
- **89 of the phone folder's contacts carry the category** "HUAWEI P40 Pro
  (contacts synced by Link to Windows)", the sync's own tag. A move keeps it
  like every other property: dropping it is the user's call, not the tool's.
- **More than 4 concurrent requests per app and mailbox are throttled**: 429
  `ApplicationThrottled`, "Application is over its MailboxConcurrency limit".
  The move sends one request at a time and caps a call at 25 contacts, about
  four requests each.
- **A contact folder** has `id`, `displayName`, `parentFolderId` and
  `wellKnownName` (`contacts` for the default one, null otherwise), and no item
  count. `$count=true` on its `/contacts` gives one, which can exceed what the
  folder lists: the default folder counted 164 and listed 162, the phone folder
  115 and 115. What the other two are is unknown, so the folder delete refuses
  any folder whose count is not 0.
- **`DELETE /me/contactFolders/{id}` is soft**: the folder lands under
  `/me/contactFolders/deleteditems/childFolders` (not in the mail folder tree)
  and its ID still resolves; `permanentDelete` purges it from there.

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

# 4. Integration, HTTP transport (temporary config on a free port, a throwaway
#    identity enrolled and removed again, driven by X-Auth-Email alone)
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
- Token cache issues: delete `~/.outlook_mcp_token_cache.json` and re-auth. For one user of an HTTP deployment the file is `<sha256(lowercased address)>.json` under `[auth].cache_dir`, or under `~/.outlook_mcp/caches/` when unset; `outlook_mcp.auth.user_cache_path()` computes it, and `shared_cache_path()` the single-account one. Never guess the path: ask the deployed code, `load_config(<its toml>)` then `user_cache_path(address, config.cache_directory)`
- "\<user\> has not authorized this server": that user has no cache, or nothing usable in it. Not a bug, and deliberately not a fallback: they enrol at `/oauth/login`, or an operator runs `outlook-mcp-auth --user <them>`
- "The authorization ... does not include Contacts.ReadWrite": the user has not consented to the scope, or this installation's token has not caught up with a consent given elsewhere. Every other tool still works; the fix is one more sign-in on that installation (`python outlook_mcp_auth.py`, or `/oauth/login` for a user of an HTTP deployment), and the running server adopts the rewritten cache on the next call. Never copy a cache in from another host instead (see Deployment)
- A sign-in that "did not take": the server logs `Token cache <file> was rewritten, reloaded` when it adopts one. No such line means the sign-in wrote a different file than the server reads; compute both with `user_cache_path()` / `shared_cache_path()` rather than guessing
- "No valid token available", or "Could not obtain a token for this client id", together with `AADSTS700016` from MSAL, means the app registration behind the client id no longer exists in the directory: that is an Azure-side problem, not a code regression. The second message comes from the secret verification described above, which is the first thing to fail when the registration is gone
- A download link that answers 404: it is single use and the fetch deleted the file, or it expired after `downloads.TICKET_TTL_SECONDS`. The table is in memory, so a restart invalidates every outstanding link. All of that is by design; the fix is to call `outlook_get_attachment` again. A **401 or a redirect to a login page** is a different problem and not the server's: the proxy is authenticating `/attachments/`, which it must not
- A downloaded file that vanished from `download_path` before anyone deleted it: `[attachments].retention_minutes` (default 60). The sweep logs how many it took
- Graph API errors: check response body in exception (includes error code and message)
- Rate limiting: Graph returns 429 with Retry-After header (not auto-handled currently)
- The server must never depend on its working directory: a stdio server inherits the CWD of whatever host spawned it, which may not even be traversable by the server's user (e.g. a daemon started from another user's 0700 home). With mcp SDK 1.x this used to crash at startup because FastMCP's pydantic-settings probed `./.env`; SDK 2.x reads no `.env` / `MCP_*` at all, and the one path this project resolves itself (`outlook_mcp.toml` via `config.DEFAULT_CONFIG_PATH`) hangs off `config.PROJECT_ROOT`, which is `__file__`-derived. Keep it that way: no relative path, no `Path.cwd()`

## Microsoft Graph API Quirks

- **OData queries** ($filter, $select, $orderBy) have strict syntax - check Graph docs
- **Pagination** uses `@odata.nextLink` (not implemented in tools - uses $top/$skip instead)
- **Recurrence expansion** for calendar events requires `startDateTime` and `endDateTime` query params
- **Meeting creation** sets `isOnlineMeeting: true` to auto-generate Teams link
- **Folder moves** accept either folder ID or well-known name string
- **Attendee types** are: `required`, `optional`, `resource`
- **Contacts**: `/me/contacts` is the default folder only; `$filter` reaches only `emailAddresses/any(a:a/address eq ...)`; a PATCH without `displayName` may regenerate it; see "Contacts" above
- **Contacts cannot be moved or copied**, and a copy with a photo does not survive its original's delete; see "Moving Contacts, and Contact Folders"
- **Concurrency**: more than 4 simultaneous requests per app and mailbox get 429 `ApplicationThrottled`
