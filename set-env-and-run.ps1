# PowerShell script to set Azure environment variables and run the OneNote MCP Server (Delegated Auth)
# Usage: .\set-env-and-run.ps1
# NOTE: This version uses delegated authentication - no client secret required

Write-Host "Setting Azure environment variables and starting server (Delegated Auth)..." -ForegroundColor Green

# Check if secret.local file exists
if (Test-Path "secret.local") {
    Write-Host "Loading secrets from secret.local file..." -ForegroundColor Cyan
    
    # Read the secret.local file and parse KEY=VALUE or KEY: VALUE pairs
    $secretLines = Get-Content "secret.local" | Where-Object { $_ -notmatch "^#" -and $_ -ne "" }
    
    foreach ($line in $secretLines) {
        # Support both KEY=VALUE and KEY: VALUE formats
        if ($line -match "^(.+?)[:=]\s*(.+)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            
            switch ($key) {
                "AZURE_TENANT_ID" { $env:AZURE_TENANT_ID = $value }
                "AZURE_CLIENT_ID" { $env:AZURE_CLIENT_ID = $value }
                # AZURE_CLIENT_SECRET not needed for delegated authentication
            }
        }
    }
} else {
    Write-Host "secret.local file not found. Using placeholder values." -ForegroundColor Yellow
    Write-Host "Create secret.local file with your actual credentials (see secret.local.example)" -ForegroundColor Yellow
    
    $env:AZURE_TENANT_ID = "<YOUR_TENANT_ID>"
    $env:AZURE_CLIENT_ID = "<YOUR_CLIENT_ID>"
    # AZURE_CLIENT_SECRET not needed for delegated authentication
}

Write-Host "Environment variables set successfully!" -ForegroundColor Green
Write-Host "AZURE_TENANT_ID: $env:AZURE_TENANT_ID" -ForegroundColor Yellow
Write-Host "AZURE_CLIENT_ID: $env:AZURE_CLIENT_ID" -ForegroundColor Yellow
Write-Host "Note: Using delegated authentication - no client secret required" -ForegroundColor Cyan

Write-Host "`nBuilding and starting the server..." -ForegroundColor Cyan

# Build the project
& npm run build

# Start the server
& npm run start
