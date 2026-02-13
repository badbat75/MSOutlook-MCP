#!/usr/bin/env bash
# Outlook MCP - Environment Setup Script
# ========================================
# This script loads environment variables from .env file and activates venv
#
# Usage:
#   1. Copy .env.example to .env
#   2. Edit .env and fill in your Azure AD credentials
#   3. Run: source ./scripts/setup-env.sh
#
# Note: Use 'source' to run in the current shell session

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

echo -e "${CYAN}================================================================${NC}"
echo -e "${CYAN}Outlook MCP - Environment Setup${NC}"
echo -e "${CYAN}================================================================${NC}"
echo ""

# =============================================================================
# Load .env file
# =============================================================================

ENV_FILE="$PROJECT_ROOT/.env"

if [ ! -f "$ENV_FILE" ]; then
    echo -e "${RED}ERROR: .env file not found!${NC}"
    echo ""
    echo -e "${YELLOW}Please create a .env file:${NC}"
    echo -e "${WHITE}  1. Copy .env.example to .env${NC}"
    echo -e "     ${GRAY}cp .env.example .env${NC}"
    echo -e "${WHITE}  2. Edit .env and fill in your Azure AD credentials${NC}"
    echo ""
    echo -e "${CYAN}Get credentials from:${NC}"
    echo -e "${WHITE}  https://entra.microsoft.com > App registrations${NC}"
    echo ""
    return 1 2>/dev/null || exit 1
fi

echo -e "${GRAY}Loading configuration from .env...${NC}"

# Parse .env file and set environment variables
while IFS= read -r line || [ -n "$line" ]; do
    # Trim whitespace
    line=$(echo "$line" | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')

    # Skip empty lines and comments
    if [ -z "$line" ] || [[ "$line" =~ ^# ]]; then
        continue
    fi

    # Parse KEY=VALUE format
    if [[ "$line" =~ ^([^=]+)=(.*)$ ]]; then
        key="${BASH_REMATCH[1]}"
        value="${BASH_REMATCH[2]}"

        # Trim whitespace from key and value
        key=$(echo "$key" | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')
        value=$(echo "$value" | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')

        # Remove quotes if present
        value=$(echo "$value" | sed -e 's/^["'"'"']//' -e 's/["'"'"']$//')

        # Set environment variable
        export "${key}=${value}"
    fi
done < "$ENV_FILE"

# =============================================================================
# Validation
# =============================================================================

MISSING_VARS=()
EMPTY_VARS=()

REQUIRED_VARS=(
    "OUTLOOK_CLIENT_ID"
    "OUTLOOK_CLIENT_SECRET"
    "OUTLOOK_TENANT_ID"
)

for var in "${REQUIRED_VARS[@]}"; do
    value="${!var}"

    if [ -z "$value" ]; then
        EMPTY_VARS+=("$var")
    fi
done

if [ ${#MISSING_VARS[@]} -gt 0 ] || [ ${#EMPTY_VARS[@]} -gt 0 ]; then
    echo -e "${RED}ERROR: Missing or empty environment variables!${NC}"
    echo ""

    if [ ${#MISSING_VARS[@]} -gt 0 ]; then
        echo -e "${YELLOW}Missing variables:${NC}"
        for var in "${MISSING_VARS[@]}"; do
            echo -e "${WHITE}  - $var${NC}"
        done
    fi

    if [ ${#EMPTY_VARS[@]} -gt 0 ]; then
        echo -e "${YELLOW}Empty variables:${NC}"
        for var in "${EMPTY_VARS[@]}"; do
            echo -e "${WHITE}  - $var${NC}"
        done
    fi

    echo ""
    echo -e "${CYAN}Please edit your .env file and set these values.${NC}"
    echo -e "${WHITE}Get credentials from: https://entra.microsoft.com${NC}"
    echo ""
    return 1 2>/dev/null || exit 1
fi

# =============================================================================
# Activate Virtual Environment
# =============================================================================

VENV_PATH="$PROJECT_ROOT/venv/bin/activate"

if [ -f "$VENV_PATH" ]; then
    echo -e "${GREEN}Activating virtual environment...${NC}"
    source "$VENV_PATH"
    echo -e "${GREEN}Virtual environment activated${NC}${DARK_GRAY} (venv)${NC}"
else
    echo -e "${YELLOW}WARNING: Virtual environment not found${NC}"
    echo -e "${WHITE}Run: python -m venv venv${NC}"
    echo ""
fi

# =============================================================================
# Display Configuration
# =============================================================================

echo ""
echo -e "${GREEN}Environment configured:${NC}"

CLIENT_ID="${OUTLOOK_CLIENT_ID}"
TENANT_ID="${OUTLOOK_TENANT_ID}"

echo -ne "${WHITE}  OUTLOOK_CLIENT_ID     = ${NC}"
if [ ${#CLIENT_ID} -gt 8 ]; then
    echo -e "${GRAY}${CLIENT_ID:0:8}...${NC}"
else
    echo -e "${GRAY}${CLIENT_ID}${NC}"
fi

echo -ne "${WHITE}  OUTLOOK_CLIENT_SECRET = ${NC}"
echo -e "${GRAY}***${DARK_GRAY} (hidden)${NC}"

echo -ne "${WHITE}  OUTLOOK_TENANT_ID     = ${NC}"
echo -e "${GRAY}${TENANT_ID}${NC}"

echo ""
echo -e "${CYAN}You can now run:${NC}"
echo -e "${WHITE}  python outlook_mcp_auth.py       ${DARK_GRAY}# Initial authorization${NC}"
echo -e "${WHITE}  python outlook_mcp_server.py     ${DARK_GRAY}# Start MCP server${NC}"
echo ""
