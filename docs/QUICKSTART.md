# Quick Start Guide

Get up and running with Outlook MCP Server in 5 minutes.

## Prerequisites

- Python 3.11 or higher
- Azure AD App registration (see main README.md for details)
- Your Azure App credentials (Client ID, Client Secret, Tenant ID)

## Setup Steps

### 1. Set Up Virtual Environment

```powershell
# Create and activate virtual environment (already done if following setup)
python -m venv venv
venv\Scripts\activate

# Install dependencies (already done if following setup)
pip install -r requirements.txt
```

### 2. Configure Environment Variables

**Using .env file (Recommended)**

**Windows (PowerShell):**
```powershell
# 1. Create your config file from the template
Copy-Item .env.example .env

# 2. Edit .env and fill in your credentials
#    OUTLOOK_CLIENT_ID=your-actual-client-id
#    OUTLOOK_CLIENT_SECRET=your-actual-client-secret
#    OUTLOOK_TENANT_ID=your-actual-tenant-id

# 3. Load the environment
. .\scripts\setup-env.ps1
```

**macOS/Linux (Bash):**
```bash
# 1. Create your config file from the template
cp .env.example .env

# 2. Edit .env and fill in your credentials
#    OUTLOOK_CLIENT_ID=your-actual-client-id
#    OUTLOOK_CLIENT_SECRET=your-actual-client-secret
#    OUTLOOK_TENANT_ID=your-actual-tenant-id

# 3. Load the environment
source ./scripts/setup-env.sh
```

**Manual Setup (Session Only)**

**Windows:**
```powershell
$env:OUTLOOK_CLIENT_ID = "your-client-id"
$env:OUTLOOK_CLIENT_SECRET = "your-client-secret"
$env:OUTLOOK_TENANT_ID = "your-tenant-id"
```

**macOS/Linux:**
```bash
export OUTLOOK_CLIENT_ID="your-client-id"
export OUTLOOK_CLIENT_SECRET="your-client-secret"
export OUTLOOK_TENANT_ID="your-tenant-id"
```

### 3. Authorize the Application

```powershell
python outlook_mcp_auth.py
```

This will:
- Open your browser for Microsoft login
- Redirect to a local callback server
- Save OAuth tokens to `~/.outlook_mcp_token_cache.json`

### 4. Start the Server

```powershell
# For Claude Desktop (stdio mode)
python outlook_mcp_server.py

# For remote access (HTTP mode)
python outlook_mcp_server.py --http --port 8000
```

### 5. Configure Claude Desktop

Use the config generator script to automatically create the correct configuration with your paths and credentials from `.env`:

**Windows:**
```powershell
# Preview the generated config
.\scripts\generate-claude-config.ps1

# Install directly into Claude Desktop config (recommended)
.\scripts\generate-claude-config.ps1 -Install

# Or write to a custom file
.\scripts\generate-claude-config.ps1 -OutFile .\my-config.json
```

**macOS/Linux:**
```bash
# Preview the generated config
./scripts/generate-claude-config.sh

# Install directly into Claude Desktop config (recommended)
./scripts/generate-claude-config.sh --install

# Or write to a custom file
./scripts/generate-claude-config.sh --outfile ./my-config.json
```

The install flag will:
- Auto-detect your venv Python path and server script location
- Load credentials from your `.env` file
- Merge with any existing Claude Desktop config (preserving other MCP servers)
- Write to the appropriate config location for your OS

Restart Claude Desktop and you're ready to go!

## Daily Usage

Each time you want to work on this project:

**Windows:**
```powershell
# Load environment from .env and activate venv in one command
. .\scripts\setup-env.ps1

# Now you can run any commands
python outlook_mcp_auth.py      # If you need to re-authorize
python outlook_mcp_server.py    # Start the server
```

**macOS/Linux:**
```bash
# Load environment from .env and activate venv in one command
source ./scripts/setup-env.sh

# Now you can run any commands
python outlook_mcp_auth.py      # If you need to re-authorize
python outlook_mcp_server.py    # Start the server
```

## Testing

Ask Claude to:
- "Show me my unread emails"
- "What meetings do I have today?"
- "Send an email to someone@example.com"
- "Create a meeting for tomorrow at 2pm"

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `401 Unauthorized` | Run `python outlook_mcp_auth.py` again |
| `403 Forbidden` | Check API permissions in Azure AD |
| `ModuleNotFoundError` | Activate venv: `venv\Scripts\activate` (Windows) or `source venv/bin/activate` (macOS/Linux) |
| Environment variables not set | Run `. .\scripts\setup-env.ps1` (Windows) or `source ./scripts/setup-env.sh` (macOS/Linux) |

## Next Steps

- Read [CLAUDE.md](../CLAUDE.md) for architectural details
- Read [README.md](../README.md) for full documentation
- Check the tool descriptions in `outlook_mcp/server.py`
