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

## 3. Configure the server

```bash
# Windows
Copy-Item outlook_mcp.toml.example outlook_mcp.toml
# macOS/Linux
cp outlook_mcp.toml.example outlook_mcp.toml
```

Fill in the values from step 1:

```toml
[credentials]
client_id = "your-client-id"
client_secret = "your-client-secret"
tenant_id = "common"          # or your directory (tenant) ID
```

**This file is the only place the server is configured.** It reads nothing else
from the environment, and it never takes a credential from a request. Both the
server and `outlook_mcp_auth.py` find it from the project root whatever the
working directory is, so nothing needs loading before either of them starts.

Two rules worth knowing:

- **The location can be moved** with `--config PATH` or `$OUTLOOK_MCP_CONFIG`.
  A path given either way must exist; a typo is an error, not a silent
  fallback to a server with no credentials.
- **Invalid contents are a hard error too**, with exit code 2 and a message
  naming the file and the key.

`outlook_mcp.toml` is gitignored. Never commit it.

---

## 4. Authorize (once)

```bash
python outlook_mcp_auth.py                # opens a browser
python outlook_mcp_auth.py --no-browser   # headless/SSH: prints the URL
python outlook_mcp_auth.py --code 'http://localhost:5000/callback?code=...'
outlook-mcp-auth                          # the installed console script
```

Sign in, accept the permissions, and the tokens are written to
`~/.outlook_mcp_token_cache.json`. MSAL refreshes them from then on; re-run
this only if you see `401` errors that persist.

If the browser opens but the callback never arrives (a remote machine, a
firewall), press Ctrl+C and paste the full callback URL when prompted.

Serving several people from one server is a different story: each of them gets
a token cache of their own, and they authorize themselves. See
[Multi-user behind a reverse proxy](#multi-user-behind-a-reverse-proxy).

---

## 5. Choose the transport

The transport and the HTTP listening address live in the same
`outlook_mcp.toml`, never on the command line, so the server behaves the same
whoever starts it. **The default is stdio**, which is what Claude Desktop and
Claude Code want, so most setups can leave `[server]` alone.

```toml
[server]
transport = "http"        # "stdio" (default) or "http"
bind_host = "127.0.0.1"   # HTTP only, and loopback only
bind_port = 8000          # HTTP only
```

A file that exists but cannot be parsed, or that names an unknown transport or
an out-of-range port, stops the server with exit code 2. It never falls back to
stdio silently: an HTTP deployment that quietly listens on nothing is worse
than one that refuses to start.

**`bind_host` must be a loopback address.** Over HTTP a request carries no
credential at all: it names the mailbox it is addressing, in a header a reverse
proxy appends after authenticating the user. Bound anywhere else, anyone able
to reach the port could name any mailbox, so the server refuses to start.

---

## 6. Start the server

```bash
python outlook_mcp_server.py                       # transport from outlook_mcp.toml
python outlook_mcp_server.py --config /etc/outlook_mcp.toml
outlook-mcp                                        # the installed console script
```

The three forms are the same program.

The three forms are the same program, and both transports run as the one app
registration in `[credentials]`. What changes is whose mailbox a call reaches:

| Transport | Whose mailbox | Token cache | Endpoint |
|-----------|---------------|-------------|----------|
| `stdio` | the one account that authorized in step 4 | `~/.outlook_mcp_token_cache.json` | process stdin/stdout |
| `http` | the user named in `X-Auth-Email`, appended by the reverse proxy | `~/.outlook_mcp/caches/<hash>.json`, one per user | `http://127.0.0.1:<bind_port>/mcp`, behind the proxy |

On stdio there is no authentication and none is possible: whoever can start the
process already reads the configuration file and the token cache, so the
operating system user is the boundary. That is fine when each person runs their
own copy, and it is exactly what stops working when one daemon spawns the
server for several people; the HTTP deployment is the answer to that case.

Nothing a caller sends is ever a credential, in either mode, and the download
path stays server side: a remote caller must not get to choose where the server
writes files.

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

Use the venv interpreter, not a bare `python`. No `env` block is needed: the
server reads `outlook_mcp.toml` from the project root. Point a particular host
at a different file with `"args": [..., "--config", "C:\\path\\to\\other.toml"]`.

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

# HTTP, remote server (the proxy authenticates you and adds the identity header)
claude mcp add --transport http outlook https://outlook-mcp.example.com/mcp
```

The HTTP client sends no credential of its own. It authenticates to the reverse
proxy however that proxy is configured, and the proxy is what tells the server
whose mailbox to open. Any MCP client that supports streamable HTTP works the
same way.

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

## Multi-user behind a reverse proxy

One server, several mailboxes, no credential on the wire. The reverse proxy
authenticates the user however you already authenticate people, and appends
their address; the server keeps one MSAL token cache per user and never lets
one person's grant answer another person's request.

**The whole arrangement rests on one property: the proxy must be the only way
in.** The identity header proves nothing by itself, it is believable only
because the proxy replaces whatever the client sent. That is why the server
refuses to start on anything but a loopback bind, and why the proxy must set
the header on every location it forwards from.

### 1. Bind to loopback and name the header

```toml
[server]
transport = "http"
bind_host = "127.0.0.1"
bind_port = 8000
allowed_hosts = ["outlook-mcp.example.com"]       # the Host the proxy forwards

[credentials]
client_id = "..."
client_secret = "..."

[auth]
user_header = "X-Auth-Email"
public_url = "https://outlook-mcp.example.com"    # turns on browser enrollment
```

**`allowed_hosts` is not optional behind a proxy.** A loopback bind makes the
MCP SDK turn on DNS-rebinding protection by itself, after which it accepts a
Host header of localhost and nothing else, and every proxied request comes back
**421 Invalid Host header**. That check assumes a proxied server binds a
routable address; this one binds loopback on purpose, which is the case the rule
does not anticipate. Naming the site here widens the list and keeps the
protection on. Write Host header values, not URLs: `outlook-mcp.example.com`,
never `https://outlook-mcp.example.com`.

### 2. Put NGINX in front of it

```nginx
server {
    listen 443 ssl;
    server_name outlook-mcp.example.com;

    # However you authenticate people. auth_request is one way; the point is
    # that $authenticated_user below is set by you, never by the client.
    auth_request      /_auth;
    auth_request_set  $authenticated_user $upstream_http_x_auth_request_email;

    location / {
        proxy_pass http://127.0.0.1:8000;

        # This line is the security boundary. proxy_set_header REPLACES the
        # client's version of the header, so a caller cannot smuggle in an
        # address of its own. It has to be in EVERY location that reaches the
        # server: /mcp, /oauth/ and /attachments/ alike.
        proxy_set_header X-Auth-Email $authenticated_user;

        # Streamable HTTP holds long responses open.
        proxy_buffering off;
        proxy_read_timeout 300s;
    }
}
```

A location that forwards without that `proxy_set_header` line passes the
client's own header through, and the whole scheme collapses. The server refuses
a request whose header is missing or blank, so a forgotten line fails closed
rather than silently serving the wrong mailbox.

### 3. Enrol each user

Either they do it themselves, once, in a browser:

```
https://outlook-mcp.example.com/oauth/login
```

They arrive already authenticated by the proxy, sign in to Microsoft, and the
tokens land in their own cache. Register `<public_url>/oauth/callback` as a
redirect URI on the app registration first, alongside the localhost one from
step 1.

Or an operator does it for them on the server:

```bash
outlook-mcp-auth --user someone@example.com
```

Either way, enrolling again replaces that user's previous grant rather than
adding a second account to the same cache. A user who has never enrolled gets
an error saying so; they are never quietly served an application-level token.

### 4. Getting attachments back out

`outlook_get_attachment` writes the file to the machine the **server** runs on.
Over stdio that is also the caller's machine, so the path it answers with is
one they can open. Over HTTP it is not, so with `public_url` set the tool
answers with a link instead:

```
https://outlook-mcp.example.com/attachments/<token>
```

The token is minted by the server, and it is good for **one fetch, for fifteen
minutes, by the user it was issued to**. The route runs the same identity check
the tools do, so a link copied into a shared transcript is worth nothing to the
next reader, and nothing at all once it has been used. Which means `/attachments/`
needs the same `proxy_set_header` line as everything else the proxy forwards.

Files are written per user and per message:

```
<download_path>/<sha256 of the address>/msg-<hash of the message id>/invoice.pdf
```

One person's downloads never appear in another's directory, which is the same
isolation their token caches have. The per-message level is what lets a caller
clean up after itself without ever naming a path:

```
outlook_delete_attachment_files(message_id="AAMk...")            # every file
outlook_delete_attachment_files(message_id="AAMk...", filename="invoice.pdf")
```

That deletes from the server's filesystem only. The attachment stays in the
mailbox and can be downloaded again. Nothing expires the files on its own, so
on a busy deployment either tell the agent to clean up, or age
`<download_path>` out with a systemd timer or `tmpfiles.d`.

---

## Deploying to a Linux host

`scripts/deploy.sh` does the whole of the above on a remote machine: it copies
the checkout, builds a virtual environment, installs the configuration and a
systemd unit, and starts the service. Run it from the checkout (on Windows,
from Git Bash); it needs `ssh`, `scp`, `git` and `tar`.

```bash
./scripts/deploy.sh --target root@server.example.com

./scripts/deploy.sh --target me@host \
    --dir /srv/outlook-mcp \
    --service-user outlook \
    --service-name outlook-mcp \
    --config deploy/prod.toml
```

| Option | Default | What it is |
|--------|---------|------------|
| `--target USER@HOST` | required | SSH destination, root or able to `sudo` |
| `--dir PATH` | `/opt/outlook-mcp` | Where everything is installed |
| `--data-dir PATH` | `<dir>/data` | Token caches and attachments. It is the service's `HOME`, and the one path the unit lets it write |
| `--service-user NAME` | `outlook-mcp` | The system user the service runs as, created if missing |
| `--service-name NAME` | `outlook-mcp` | The unit name, without `.service` |
| `--config PATH` | none | A local TOML to install as the server configuration |
| `--port N` | `8000` | The port in the starter configuration |
| `--python PATH` | `python3` | The interpreter used to build the venv on the host |
| `--no-restart` | | Install everything, leave the running service alone |
| `--dry-run` | | Print the script that would run on the host, and stop |

What lands on the host:

```
/opt/outlook-mcp/
├── app/                 the checkout, replaced whole on every deploy
├── venv/                built once, reused
├── outlook_mcp.toml     the configuration, 0600, owned by the service user
└── data/                everything the service writes, 0700
    ├── caches/          one MSAL token cache per enrolled mailbox
    └── attachments/     what the download tools write
```

The configuration names both those directories (`[auth].cache_dir` and
`[attachments].download_path`), so nothing depends on where the service user's
home happens to be. The unit still sets `HOME` to the data directory, but only
so that a stray `~` resolves inside it rather than in a shared account's home.

One tree to back up, and one to delete. The configuration and the data sit
**beside** the code rather than inside it, so a deploy can replace `app/`
wholesale (no file survives a rename) without touching either the client secret
or anybody's authorization.

Five things the script does that are worth knowing:

- **Only what git tracks is uploaded.** Your local `outlook_mcp.toml` and `.env`
  are gitignored, so they never travel by accident. The configuration the
  service runs on is the one from `--config`, or one already on the host, or a
  starter file the script writes for you to fill in.
- **The configuration is validated on the host before the service is
  restarted**, by the server's own loader. A non-loopback bind or an unknown
  transport fails the deploy with the real error instead of leaving a unit
  crash-looping every five seconds.
- **`transport = "http"` is required.** A stdio server is spawned once per
  client by whatever MCP host talks to it; under systemd it would exit
  immediately, so the script refuses rather than installing something that
  cannot work.
- **A client secret never sits in `/tmp`.** `--config` is uploaded into a 0700
  staging directory in the SSH user's home, installed 0600, and deleted there.
- **An attachment path outside the data directory fails the deploy.** The unit
  grants exactly one writable path, so anywhere else would only break when
  somebody downloads their first file, a long way from the cause.

### Enrolling users on the host

Pass the same `--config` the service runs on. That file is what says where the
caches live, so the tokens land where the service will look for them.

```bash
sudo -u outlook-mcp /opt/outlook-mcp/venv/bin/outlook-mcp-auth \
     --config /opt/outlook-mcp/outlook_mcp.toml \
     --user them@example.com --no-browser
```

It prints a URL: open it in any browser, sign in, and paste the whole callback
URL back at the prompt. The browser will fail to load that callback, which does
not matter; the code is in the address bar.

Setting `[auth].public_url` and letting each person use `/oauth/login` avoids
this entirely.

**Give each installation its own sign-in.** Copying a token cache from another
machine looks like a shortcut and is not one: Entra rotates a refresh token as
it redeems it, so two copies redeeming independently keep invalidating each
other, and the failure arrives later and looks random.

### Afterwards

```bash
systemctl status outlook-mcp
journalctl -u outlook-mcp -f
```

The unit runs with `ProtectSystem=strict`, an empty capability set and a
`@system-service` syscall filter. Its one writable directory is the data
directory, which it also runs with as `HOME`: `ProtectHome` is therefore off,
and turning it on breaks every mailbox with an error that points nowhere near
the cause.

`RestartPreventExitStatus=2` is deliberate. Exit code 2 is the server's
configuration error, and no number of restarts will fix an unreadable TOML, so
the unit fails visibly instead.

An **stdio** deployment needs none of this: no unit, no service user, just the
checkout, a config file and an MCP host that spawns the process. See
[contrib/openclaw.json](../contrib/openclaw.json).

---

## Daily usage

The server reads `outlook_mcp.toml` on its own, so nothing needs loading before
it starts, from any shell or any working directory.

---

## Troubleshooting

| Symptom | Cause and fix |
|---------|---------------|
| `ModuleNotFoundError` | The venv is not active: `venv\Scripts\activate` |
| `401 Unauthorized` | Token cache stale or missing: re-run `python outlook_mcp_auth.py` |
| `403 Forbidden` | A delegated permission is missing in the app registration (step 1.7) |
| `429` rate limited | Wait the seconds named in the message and retry |
| `Configuration error:` and exit code 2 | `outlook_mcp.toml` is unreadable or invalid; the message names the file and the key |
| Refuses to start: `bind_host must be a loopback address` | An HTTP server bound where clients could reach it directly. Bind `127.0.0.1` and let the reverse proxy be the only way in |
| `No X-Auth-Email header on this request` | A proxy location forwarding without `proxy_set_header X-Auth-Email` |
| `421 Invalid Host header` | From the MCP SDK, not nginx: a loopback bind turns on DNS-rebinding protection, which accepts only localhost. Put the site in `[server].allowed_hosts` |
| `<user> has not authorized this server` | That user has never enrolled: `/oauth/login`, or `outlook-mcp-auth --user <them>` |
| An attachment answers with a server path instead of a link | `[auth].public_url` is not set, so the server cannot say what URL it is reachable on |
| `This download link is not valid` | It has been used, it is over fifteen minutes old, or it was issued to somebody else. Ask for the attachment again |
| Browser callback never arrives | Ctrl+C, then paste the callback URL by hand, or use `--no-browser` |
| `AADSTS...` errors from Microsoft | See the error-by-error table in [SETUP_PERSONAL_ACCOUNTS.md](SETUP_PERSONAL_ACCOUNTS.md#troubleshooting) |
