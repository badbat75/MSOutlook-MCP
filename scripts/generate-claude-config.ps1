# Outlook MCP - Claude Desktop Config Generator
# ===============================================
# Generates claude_desktop_config.json with the correct absolute paths.
#
# No credentials are written into it: the server reads outlook_mcp.toml itself,
# from the project root, whatever directory Claude Desktop spawns it in.
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
# Warn if the server has nothing to run as
# =============================================================================

$configFile = Join-Path $projectRoot "outlook_mcp.toml"

if (-not (Test-Path $configFile)) {
    Write-Host "WARNING: outlook_mcp.toml not found." -ForegroundColor Yellow
    Write-Host "  The generated config is still correct, but the server will have" -ForegroundColor Gray
    Write-Host "  no credentials until you create it:" -ForegroundColor Gray
    Write-Host "    Copy-Item outlook_mcp.toml.example outlook_mcp.toml" -ForegroundColor White
    Write-Host ""
}

# =============================================================================
# Build config object
# =============================================================================

# No "env" block: the server reads outlook_mcp.toml itself. Point a particular
# host at a different file by adding "--config", "<path>" to args.
$config = @{
    mcpServers = @{
        MS_Outlook_MCP = @{
            command = $venvPython
            args    = @($serverScript)
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
Write-Host "  Config:  $configFile" -ForegroundColor DarkGray
Write-Host ""
