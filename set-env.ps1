# PowerShell script to set Azure environment variables for OneNote MCP Server
# Usage: .\set-env.ps1

Write-Host "Setting Azure environment variables..." -ForegroundColor Green

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
                # Note: Still parsing AZURE_CLIENT_SECRET for backward compatibility
                "AZURE_CLIENT_SECRET" { $env:AZURE_CLIENT_SECRET = $value }
            }
        }
    }
} else {
    Write-Host "secret.local file not found. Using placeholder values." -ForegroundColor Yellow
    Write-Host "Create secret.local file with your actual credentials (see secret.local.example)" -ForegroundColor Yellow
    
    $env:AZURE_TENANT_ID = "<YOUR_TENANT_ID>"
    $env:AZURE_CLIENT_ID = "<YOUR_CLIENT_ID>"
    $env:AZURE_CLIENT_SECRET = "<YOUR_CLIENT_SECRET>"
}

Write-Host "Environment variables set successfully!" -ForegroundColor Green
Write-Host "AZURE_TENANT_ID: $env:AZURE_TENANT_ID" -ForegroundColor Yellow
Write-Host "AZURE_CLIENT_ID: $env:AZURE_CLIENT_ID" -ForegroundColor Yellow
Write-Host "AZURE_CLIENT_SECRET: $env:AZURE_CLIENT_SECRET" -ForegroundColor Yellow

Write-Host "`nYou can now run the server with: npm start" -ForegroundColor Cyan
