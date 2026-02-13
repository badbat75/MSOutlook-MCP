# Outlook MCP - Environment Setup Script
# ========================================
# This script loads environment variables from .env file and activates venv
#
# Usage:
#   1. Copy .env.example to .env
#   2. Edit .env and fill in your Azure AD credentials
#   3. Run: . .\scripts\setup-env.ps1
#
# Note: Use dot-sourcing (. .\scripts\setup-env.ps1) to run in the current session

$projectRoot = Split-Path $PSScriptRoot -Parent

Write-Host "=" -NoNewline -ForegroundColor Cyan
Write-Host ("=" * 58) -ForegroundColor Cyan
Write-Host "Outlook MCP - Environment Setup" -ForegroundColor Cyan
Write-Host "=" -NoNewline -ForegroundColor Cyan
Write-Host ("=" * 58) -ForegroundColor Cyan
Write-Host ""

# =============================================================================
# Load .env file
# =============================================================================

$envFile = Join-Path $projectRoot ".env"

if (-not (Test-Path $envFile)) {
    Write-Host "ERROR: .env file not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please create a .env file:" -ForegroundColor Yellow
    Write-Host "  1. Copy .env.example to .env" -ForegroundColor White
    Write-Host "     " -NoNewline
    Write-Host "Copy-Item .env.example .env" -ForegroundColor Gray
    Write-Host "  2. Edit .env and fill in your Azure AD credentials" -ForegroundColor White
    Write-Host ""
    Write-Host "Get credentials from:" -ForegroundColor Cyan
    Write-Host "  https://entra.microsoft.com > App registrations" -ForegroundColor White
    Write-Host ""
    return
}

Write-Host "Loading configuration from .env..." -ForegroundColor Gray

# Parse .env file and set environment variables
Get-Content $envFile | ForEach-Object {
    $line = $_.Trim()

    # Skip empty lines and comments
    if ($line -eq "" -or $line.StartsWith("#")) {
        return
    }

    # Parse KEY=VALUE format
    if ($line -match "^([^=]+)=(.*)$") {
        $key = $matches[1].Trim()
        $value = $matches[2].Trim()

        # Remove quotes if present
        $value = $value -replace '^["'']|["'']$', ''

        # Set environment variable
        Set-Item -Path "env:$key" -Value $value
    }
}

# =============================================================================
# Validation
# =============================================================================

$missingVars = @()
$emptyVars = @()

$requiredVars = @(
    "OUTLOOK_CLIENT_ID",
    "OUTLOOK_CLIENT_SECRET",
    "OUTLOOK_TENANT_ID"
)

foreach ($var in $requiredVars) {
    $value = [Environment]::GetEnvironmentVariable($var)

    if ($null -eq $value) {
        $missingVars += $var
    }
    elseif ([string]::IsNullOrWhiteSpace($value)) {
        $emptyVars += $var
    }
}

if ($missingVars.Count -gt 0 -or $emptyVars.Count -gt 0) {
    Write-Host "ERROR: Missing or empty environment variables!" -ForegroundColor Red
    Write-Host ""

    if ($missingVars.Count -gt 0) {
        Write-Host "Missing variables:" -ForegroundColor Yellow
        foreach ($var in $missingVars) {
            Write-Host "  - $var" -ForegroundColor White
        }
    }

    if ($emptyVars.Count -gt 0) {
        Write-Host "Empty variables:" -ForegroundColor Yellow
        foreach ($var in $emptyVars) {
            Write-Host "  - $var" -ForegroundColor White
        }
    }

    Write-Host ""
    Write-Host "Please edit your .env file and set these values." -ForegroundColor Cyan
    Write-Host "Get credentials from: https://entra.microsoft.com" -ForegroundColor White
    Write-Host ""
    return
}

# =============================================================================
# Activate Virtual Environment
# =============================================================================

$venvPath = Join-Path $projectRoot "venv\Scripts\Activate.ps1"

if (Test-Path $venvPath) {
    Write-Host "Activating virtual environment..." -ForegroundColor Green
    & $venvPath
    Write-Host "Virtual environment activated" -ForegroundColor Green -NoNewline
    Write-Host " (venv)" -ForegroundColor DarkGray
} else {
    Write-Host "WARNING: Virtual environment not found" -ForegroundColor Yellow
    Write-Host "Run: python -m venv venv" -ForegroundColor White
    Write-Host ""
}

# =============================================================================
# Display Configuration
# =============================================================================

Write-Host ""
Write-Host "Environment configured:" -ForegroundColor Green

$clientId = $env:OUTLOOK_CLIENT_ID
$tenantId = $env:OUTLOOK_TENANT_ID

Write-Host "  OUTLOOK_CLIENT_ID     = " -NoNewline -ForegroundColor White
if ($clientId.Length -gt 8) {
    Write-Host $clientId.Substring(0, 8) -NoNewline -ForegroundColor Gray
    Write-Host "..." -ForegroundColor Gray
} else {
    Write-Host $clientId -ForegroundColor Gray
}

Write-Host "  OUTLOOK_CLIENT_SECRET = " -NoNewline -ForegroundColor White
Write-Host "***" -NoNewline -ForegroundColor Gray
Write-Host " (hidden)" -ForegroundColor DarkGray

Write-Host "  OUTLOOK_TENANT_ID     = " -NoNewline -ForegroundColor White
Write-Host $tenantId -ForegroundColor Gray

Write-Host ""
Write-Host "You can now run:" -ForegroundColor Cyan
Write-Host "  python outlook_mcp_auth.py       " -NoNewline -ForegroundColor White
Write-Host "# Initial authorization" -ForegroundColor DarkGray
Write-Host "  python outlook_mcp_server.py     " -NoNewline -ForegroundColor White
Write-Host "# Start MCP server" -ForegroundColor DarkGray
Write-Host ""
