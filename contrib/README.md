# contrib/

Deployment-specific and debugging helpers. Nothing here is imported by the
server: the files are kept because they are useful, not because the project
needs them to run. Paths inside them refer to one particular machine, so copy
and adapt rather than use as-is.

| File | What it is |
|------|------------|
| `mcp_call.py` | Minimal stdio MCP client: spawns the server, performs the handshake and calls one tool. Useful to exercise a tool without an MCP host. |
| `openclaw.json` | MCP host configuration for the `openclaw` deployment (`/home/openclaw/MCP/MSOutlook-MCP`). A template for any Linux host that spawns the server over stdio. |
| `systemd/outlook-mcp.service` | The unit for an HTTP deployment. A template: `scripts/deploy.sh` substitutes the at-sign delimited names and installs the result. |

## mcp_call.py

```bash
python contrib/mcp_call.py outlook_get_profile '{"params":{}}'
python contrib/mcp_call.py outlook_list_mail '{"params":{"folder":"inbox","top":5}}'
```

No credential is handled here: the server this spawns reads `outlook_mcp.toml`
itself, or whatever file `$OUTLOOK_MCP_CONFIG` names, since the child process
inherits the environment.

## openclaw.json

Adapt the interpreter and script paths to the target host, then merge the
`mcpServers` entry into that host's MCP configuration. The server finds
`outlook_mcp.toml` from its own location, so the command needs no wrapper and
no `env` block; a host that keeps its configuration elsewhere adds
`"--config", "/path/to/outlook_mcp.toml"` to `args`.

## systemd/outlook-mcp.service

For the HTTP transport only. A stdio server is spawned per client by whatever
MCP host talks to it (that is what `openclaw.json` sets up) and has nothing for
a service manager to supervise.

`scripts/deploy.sh` renders and installs it; see
[docs/SETUP.md](../docs/SETUP.md#deploying-to-a-linux-host). To do it by hand,
copy the file, replace every at-sign delimited name, and drop the result in
`/etc/systemd/system/`.
