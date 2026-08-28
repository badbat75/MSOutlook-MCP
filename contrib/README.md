# contrib/

Deployment-specific and debugging helpers. Nothing here is imported by the
server: the files are kept because they are useful, not because the project
needs them to run. Paths inside them refer to one particular machine, so copy
and adapt rather than use as-is.

| File | What it is |
|------|------------|
| `mcp_call.py` | Minimal stdio MCP client: spawns the server, performs the handshake and calls one tool. Useful to exercise a tool without an MCP host. |
| `openclaw.json` | MCP host configuration for the `openclaw` deployment (`/home/openclaw/MCP/MSOutlook-MCP`). A template for any Linux host that spawns the server over stdio. |

## mcp_call.py

```bash
python contrib/mcp_call.py outlook_get_profile '{"params":{}}'
python contrib/mcp_call.py outlook_list_mail '{"params":{"folder":"inbox","top":5}}'
```

Credentials come from the project `.env`, the same file the server reads.

## openclaw.json

Adapt the interpreter and script paths to the target host, then merge the
`mcpServers` entry into that host's MCP configuration.
