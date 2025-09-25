#!/bin/bash
# Shell script to set Azure environment variables for OneNote MCP Server
# Usage: source ./set-env.sh

echo "Setting Azure environment variables..."

# Check if secret.local file exists
if [ -f "secret.local" ]; then
    echo "Loading secrets from secret.local file..."
    
    # Read the secret.local file and parse KEY=VALUE pairs
    while IFS='=' read -r key value; do
        # Skip comments and empty lines
        if [[ $key =~ ^#.* ]] || [[ -z "$key" ]]; then
            continue
        fi
        
        # Remove leading/trailing whitespace
        key=$(echo "$key" | xargs)
        value=$(echo "$value" | xargs)
        
        case "$key" in
            "AZURE_TENANT_ID")
                export AZURE_TENANT_ID="$value"
                ;;
            "AZURE_CLIENT_ID")
                export AZURE_CLIENT_ID="$value"
                ;;
            "AZURE_CLIENT_SECRET")
                export AZURE_CLIENT_SECRET="$value"
                ;;
        esac
    done < secret.local
else
    echo "secret.local file not found. Using placeholder values."
    echo "Create secret.local file with your actual credentials (see secret.local.example)"
    
    export AZURE_TENANT_ID="<YOUR_TENANT_ID>"
    export AZURE_CLIENT_ID="<YOUR_CLIENT_ID>"
    export AZURE_CLIENT_SECRET="<YOUR_CLIENT_SECRET>"
fi

echo "Environment variables set successfully!"
echo "AZURE_TENANT_ID: $AZURE_TENANT_ID"
echo "AZURE_CLIENT_ID: $AZURE_CLIENT_ID"
echo "AZURE_CLIENT_SECRET: $AZURE_CLIENT_SECRET"

echo ""
echo "You can now run the server with: npm start"
