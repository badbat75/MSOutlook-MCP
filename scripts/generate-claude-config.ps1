# Outlook MCP - Claude Desktop Config Generator
# ===============================================
# Generates claude_desktop_config.json with correct paths and credentials
#
# Usage:
#   .\scripts\generate-claude-config.ps1              # Generates config to stdout
#   .\scripts\generate-claude-config.ps1 -Install     # Writes directly to Claude Desktop config
#   .\scripts\generate-claude-config.ps1 -OutFile .\my-config.json  # Writes to custom path

param(
    [switch]$Install,
    [string]$OutFile
)

$projectRoot = Split-Path $PSScriptRoot -Parent

Write-Host "=" -NoNewline -ForegroundColor Cyan
Write-Host ("=" * 58) -ForegroundColor Cyan
Write-Host "Outlook MCP - Claude Desktop Config Generator" -ForegroundColor Cyan
Write-Host "=" -NoNewline -ForegroundColor Cyan
Write-Host ("=" * 58) -ForegroundColor Cyan
Write-Host ""

# =============================================================================
# Validate venv exists
# =============================================================================

$venvPython = Join-Path $projectRoot "venv\Scripts\python.exe"

if (-not (Test-Path $venvPython)) {
    Write-Host "ERROR: Virtual environment not found!" -ForegroundColor Red
    Write-Host "  Expected: $venvPython" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Create it with:" -ForegroundColor Yellow
    Write-Host "  python -m venv venv" -ForegroundColor White
    Write-Host "  venv\Scripts\activate" -ForegroundColor White
    Write-Host "  pip install -r requirements.txt" -ForegroundColor White
    Write-Host ""
    exit 1
}

# =============================================================================
# Validate server script exists
# =============================================================================

$serverScript = Join-Path $projectRoot "outlook_mcp_server.py"

if (-not (Test-Path $serverScript)) {
    Write-Host "ERROR: outlook_mcp_server.py not found!" -ForegroundColor Red
    Write-Host "  Expected: $serverScript" -ForegroundColor Gray
    exit 1
}

# =============================================================================
# Load .env file
# =============================================================================

$envFile = Join-Path $projectRoot ".env"

if (-not (Test-Path $envFile)) {
    Write-Host "ERROR: .env file not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please create a .env file:" -ForegroundColor Yellow
    Write-Host "  Copy-Item .env.example .env" -ForegroundColor White
    Write-Host "  # Then edit .env with your Azure AD credentials" -ForegroundColor Gray
    Write-Host ""
    exit 1
}

Write-Host "Loading credentials from .env..." -ForegroundColor Gray

$envVars = @{}

Get-Content $envFile | ForEach-Object {
    $line = $_.Trim()
    if ($line -eq "" -or $line.StartsWith("#")) { return }

    if ($line -match "^([^=]+)=(.*)$") {
        $key = $matches[1].Trim()
        $value = $matches[2].Trim() -replace '^["'']|["'']$', ''
        $envVars[$key] = $value
    }
}

# Validate required vars
$requiredVars = @("OUTLOOK_CLIENT_ID", "OUTLOOK_CLIENT_SECRET", "OUTLOOK_TENANT_ID")
$missing = $requiredVars | Where-Object { -not $envVars[$_] -or [string]::IsNullOrWhiteSpace($envVars[$_]) }

if ($missing.Count -gt 0) {
    Write-Host "ERROR: Missing or empty values in .env:" -ForegroundColor Red
    foreach ($var in $missing) {
        Write-Host "  - $var" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Edit your .env file and set these values." -ForegroundColor Cyan
    exit 1
}

# =============================================================================
# Build config object
# =============================================================================

$envConfig = @{
    OUTLOOK_CLIENT_ID     = $envVars["OUTLOOK_CLIENT_ID"]
    OUTLOOK_CLIENT_SECRET = $envVars["OUTLOOK_CLIENT_SECRET"]
    OUTLOOK_TENANT_ID     = $envVars["OUTLOOK_TENANT_ID"]
}

# Add optional OUTLOOK_DOWNLOAD_PATH if present
if ($envVars["OUTLOOK_DOWNLOAD_PATH"] -and -not [string]::IsNullOrWhiteSpace($envVars["OUTLOOK_DOWNLOAD_PATH"])) {
    $envConfig["OUTLOOK_DOWNLOAD_PATH"] = $envVars["OUTLOOK_DOWNLOAD_PATH"]
}

$config = @{
    mcpServers = @{
        MS_Outlook_MCP = @{
            command = $venvPython
            args    = @($serverScript)
            env     = $envConfig
        }
    }
}

$json = $config | ConvertTo-Json -Depth 5

# =============================================================================
# Output
# =============================================================================

if ($Install) {
    $claudeConfigDir = Join-Path $env:APPDATA "Claude"
    $claudeConfigFile = Join-Path $claudeConfigDir "claude_desktop_config.json"

    # If existing config, merge instead of overwrite
    if (Test-Path $claudeConfigFile) {
        Write-Host "Existing Claude Desktop config found." -ForegroundColor Yellow
        $existing = Get-Content $claudeConfigFile -Raw | ConvertFrom-Json

        # Convert to hashtable for merging
        $existingServers = @{}
        if ($existing.mcpServers) {
            $existing.mcpServers.PSObject.Properties | ForEach-Object {
                $existingServers[$_.Name] = $_.Value
            }
        }

        # Remove old "outlook" key if present, add new "MS_Outlook_MCP"
        $existingServers.Remove("outlook")
        $existingServers["MS_Outlook_MCP"] = $config.mcpServers.MS_Outlook_MCP

        $merged = @{ mcpServers = $existingServers }
        $json = $merged | ConvertTo-Json -Depth 5

        Write-Host "Merging 'MS_Outlook_MCP' server into existing config..." -ForegroundColor Gray
    } else {
        if (-not (Test-Path $claudeConfigDir)) {
            New-Item -ItemType Directory -Path $claudeConfigDir -Force | Out-Null
        }
        Write-Host "Creating new Claude Desktop config..." -ForegroundColor Gray
    }

    $json | Set-Content -Path $claudeConfigFile -Encoding UTF8
    Write-Host ""
    Write-Host "Config written to:" -ForegroundColor Green
    Write-Host "  $claudeConfigFile" -ForegroundColor White

} elseif ($OutFile) {
    $json | Set-Content -Path $OutFile -Encoding UTF8
    Write-Host ""
    Write-Host "Config written to:" -ForegroundColor Green
    Write-Host "  $OutFile" -ForegroundColor White

} else {
    Write-Host ""
    Write-Host "Generated config:" -ForegroundColor Green
    Write-Host ""
    Write-Host $json -ForegroundColor White
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Cyan
    Write-Host "  .\scripts\generate-claude-config.ps1 -Install   " -NoNewline -ForegroundColor White
    Write-Host "# Write to Claude Desktop config" -ForegroundColor DarkGray
    Write-Host "  .\scripts\generate-claude-config.ps1 -OutFile .\out.json" -NoNewline -ForegroundColor White
    Write-Host "  # Write to file" -ForegroundColor DarkGray
}

# =============================================================================
# Summary
# =============================================================================

Write-Host ""
Write-Host "Paths used:" -ForegroundColor Gray
Write-Host "  Python:  $venvPython" -ForegroundColor DarkGray
Write-Host "  Server:  $serverScript" -ForegroundColor DarkGray
Write-Host "  Client:  " -NoNewline -ForegroundColor DarkGray
Write-Host ($envVars["OUTLOOK_CLIENT_ID"].Substring(0, [Math]::Min(8, $envVars["OUTLOOK_CLIENT_ID"].Length)) + "...") -ForegroundColor DarkGray
Write-Host "  Tenant:  $($envVars["OUTLOOK_TENANT_ID"])" -ForegroundColor DarkGray
Write-Host ""
