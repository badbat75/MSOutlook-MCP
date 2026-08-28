#!/usr/bin/env bash
# Outlook MCP - Claude Desktop Config Generator
# ===============================================
# Generates claude_desktop_config.json with the correct absolute paths.
#
# No credentials are written into it: the server reads outlook_mcp.toml itself,
# from the project root, whatever directory Claude Desktop spawns it in.
#
# Usage:
#   ./scripts/generate-claude-config.sh              # Generates config to stdout
#   ./scripts/generate-claude-config.sh --install    # Writes directly to Claude Desktop config
#   ./scripts/generate-claude-config.sh --outfile ./my-config.json  # Writes to custom path

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

# ANSI color codes
CYAN='\033[0;36m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
GREEN='\033[0;32m'
WHITE='\033[0;37m'
GRAY='\033[0;90m'
DARK_GRAY='\033[0;90m'
NC='\033[0m' # No Color

# Parse arguments
INSTALL=false
OUTFILE=""

while [[ $# -gt 0 ]]; do
    case $1 in
        --install|-i)
            INSTALL=true
            shift
            ;;
        --outfile|-o)
            OUTFILE="$2"
            shift 2
            ;;
        *)
            echo -e "${RED}Unknown option: $1${NC}"
            echo "Usage: $0 [--install] [--outfile FILE]"
            exit 1
            ;;
    esac
done

echo -e "${CYAN}================================================================${NC}"
echo -e "${CYAN}Outlook MCP - Claude Desktop Config Generator${NC}"
echo -e "${CYAN}================================================================${NC}"
echo ""

# =============================================================================
# Validate venv exists
# =============================================================================

VENV_PYTHON="$PROJECT_ROOT/venv/bin/python"

if [ ! -f "$VENV_PYTHON" ]; then
    echo -e "${RED}ERROR: Virtual environment not found!${NC}"
    echo -e "${GRAY}  Expected: $VENV_PYTHON${NC}"
    echo ""
    echo -e "${YELLOW}Create it with:${NC}"
    echo -e "${WHITE}  python -m venv venv${NC}"
    echo -e "${WHITE}  source venv/bin/activate${NC}"
    echo -e "${WHITE}  pip install -r requirements.txt${NC}"
    echo ""
    exit 1
fi

# =============================================================================
# Validate server script exists
# =============================================================================

SERVER_SCRIPT="$PROJECT_ROOT/outlook_mcp_server.py"

if [ ! -f "$SERVER_SCRIPT" ]; then
    echo -e "${RED}ERROR: outlook_mcp_server.py not found!${NC}"
    echo -e "${GRAY}  Expected: $SERVER_SCRIPT${NC}"
    exit 1
fi

# =============================================================================
# Warn if the server has nothing to run as
# =============================================================================

CONFIG_FILE="$PROJECT_ROOT/outlook_mcp.toml"

if [ ! -f "$CONFIG_FILE" ]; then
    echo -e "${YELLOW}WARNING: outlook_mcp.toml not found.${NC}"
    echo -e "${GRAY}  The generated config is still correct, but the server will have${NC}"
    echo -e "${GRAY}  no credentials until you create it:${NC}"
    echo -e "${WHITE}    cp outlook_mcp.toml.example outlook_mcp.toml${NC}"
    echo ""
fi

# =============================================================================
# Build config object
# =============================================================================

# Escape JSON strings
escape_json() {
    printf '%s' "$1" | python3 -c 'import json,sys; print(json.dumps(sys.stdin.read())[1:-1])'
}

VENV_PYTHON_ESCAPED=$(escape_json "$VENV_PYTHON")
SERVER_SCRIPT_ESCAPED=$(escape_json "$SERVER_SCRIPT")

# No "env" block: the server reads outlook_mcp.toml itself. Point a particular
# host at a different file by adding "--config", "<path>" to args.
JSON=$(cat <<EOF
{
  "mcpServers": {
    "MS_Outlook_MCP": {
      "command": "$VENV_PYTHON_ESCAPED",
      "args": [
        "$SERVER_SCRIPT_ESCAPED"
      ]
    }
  }
}
EOF
)

# =============================================================================
# Output
# =============================================================================

if $INSTALL; then
    # Determine config directory based on OS
    if [[ "$OSTYPE" == "darwin"* ]]; then
        # macOS
        CLAUDE_CONFIG_DIR="$HOME/Library/Application Support/Claude"
    else
        # Linux
        CLAUDE_CONFIG_DIR="$HOME/.config/Claude"
    fi

    CLAUDE_CONFIG_FILE="$CLAUDE_CONFIG_DIR/claude_desktop_config.json"

    # If existing config, merge instead of overwrite
    if [ -f "$CLAUDE_CONFIG_FILE" ]; then
        echo -e "${YELLOW}Existing Claude Desktop config found.${NC}"

        # Merge configs using Python
        MERGED_JSON=$(python3 <<PYTHON_SCRIPT
import json
import sys

# Read existing config
with open("$CLAUDE_CONFIG_FILE", "r") as f:
    existing = json.load(f)

# Read new config
new_config = json.loads('''$JSON''')

# Ensure mcpServers exists
if "mcpServers" not in existing:
    existing["mcpServers"] = {}

# Remove old "outlook" key if present
if "outlook" in existing["mcpServers"]:
    del existing["mcpServers"]["outlook"]

# Add new MS_Outlook_MCP
existing["mcpServers"]["MS_Outlook_MCP"] = new_config["mcpServers"]["MS_Outlook_MCP"]

# Output merged config
print(json.dumps(existing, indent=2))
PYTHON_SCRIPT
)

        echo -e "${GRAY}Merging 'MS_Outlook_MCP' server into existing config...${NC}"
        echo "$MERGED_JSON" > "$CLAUDE_CONFIG_FILE"
    else
        mkdir -p "$CLAUDE_CONFIG_DIR"
        echo -e "${GRAY}Creating new Claude Desktop config...${NC}"
        echo "$JSON" > "$CLAUDE_CONFIG_FILE"
    fi

    echo ""
    echo -e "${GREEN}Config written to:${NC}"
    echo -e "${WHITE}  $CLAUDE_CONFIG_FILE${NC}"

elif [ -n "$OUTFILE" ]; then
    echo "$JSON" > "$OUTFILE"
    echo ""
    echo -e "${GREEN}Config written to:${NC}"
    echo -e "${WHITE}  $OUTFILE${NC}"

else
    echo ""
    echo -e "${GREEN}Generated config:${NC}"
    echo ""
    echo -e "${WHITE}$JSON${NC}"
    echo ""
    echo -e "${CYAN}Usage:${NC}"
    echo -e "${WHITE}  ./scripts/generate-claude-config.sh --install     ${DARK_GRAY}# Write to Claude Desktop config${NC}"
    echo -e "${WHITE}  ./scripts/generate-claude-config.sh --outfile ./out.json  ${DARK_GRAY}# Write to file${NC}"
fi

# =============================================================================
# Summary
# =============================================================================

echo ""
echo -e "${GRAY}Paths used:${NC}"
echo -e "${DARK_GRAY}  Python:  $VENV_PYTHON${NC}"
echo -e "${DARK_GRAY}  Server:  $SERVER_SCRIPT${NC}"
echo -e "${DARK_GRAY}  Config:  $CONFIG_FILE${NC}"
echo ""
