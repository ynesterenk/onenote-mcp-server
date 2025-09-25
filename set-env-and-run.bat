@echo off
REM Batch script to set Azure environment variables and run the OneNote MCP Server
REM Usage: set-env-and-run.bat

echo Setting Azure environment variables and starting server...

REM Check if secret.local file exists
if exist "secret.local" (
    echo Loading secrets from secret.local file...
    
    REM Read secret.local file and set variables
    for /f "usebackq tokens=1,2 delims==" %%i in ("secret.local") do (
        if "%%i"=="AZURE_TENANT_ID" set AZURE_TENANT_ID=%%j
        if "%%i"=="AZURE_CLIENT_ID" set AZURE_CLIENT_ID=%%j
        if "%%i"=="AZURE_CLIENT_SECRET" set AZURE_CLIENT_SECRET=%%j
    )
) else (
    echo secret.local file not found. Using placeholder values.
    echo Create secret.local file with your actual credentials (see secret.local.example)
    
    set AZURE_TENANT_ID=<YOUR_TENANT_ID>
    set AZURE_CLIENT_ID=<YOUR_CLIENT_ID>
    set AZURE_CLIENT_SECRET=<YOUR_CLIENT_SECRET>
)

echo Environment variables set successfully!
echo AZURE_TENANT_ID: %AZURE_TENANT_ID%
echo AZURE_CLIENT_ID: %AZURE_CLIENT_ID%
echo AZURE_CLIENT_SECRET: %AZURE_CLIENT_SECRET%

echo.
echo Building and starting the server...

REM Build the project
call npm run build

REM Start the server
call npm run start
