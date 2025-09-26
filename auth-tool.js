#!/usr/bin/env node

const { ConfidentialClientApplication } = require('@azure/msal-node');
const http = require('http');
const url = require('url');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

// Configuration
const PORT = 3001;
const REDIRECT_URI = `http://localhost:${PORT}/auth/callback`;
const TOKEN_FILE = path.join(__dirname, '.auth-token');

// Get credentials from environment or prompt
const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;

if (!tenantId || !clientId) {
  console.error('❌ Error: AZURE_TENANT_ID and AZURE_CLIENT_ID must be set');
  console.error('Run: .\\set-env.ps1 first, then run this auth tool');
  process.exit(1);
}

console.log('🔐 OneNote MCP Authentication Tool');
console.log('=================================');
console.log(`Tenant ID: ${tenantId}`);
console.log(`Client ID: ${clientId}`);
console.log('=================================\\n');

// MSAL configuration
const msalConfig = {
  auth: {
    clientId: clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET || 'not-needed',
  },
};

const cca = new ConfidentialClientApplication(msalConfig);

// Microsoft Graph scopes
const graphScopes = [
  'Notes.ReadWrite',
  'User.Read',
];

// Helper function to start local web server for OAuth callback
function startAuthServer() {
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const parsedUrl = url.parse(req.url || '', true);
      
      if (parsedUrl.pathname === '/auth/callback') {
        const authCode = parsedUrl.query.code;
        const error = parsedUrl.query.error;
        
        if (error) {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end(`
            <html>
              <body style="font-family: Arial; padding: 50px; text-align: center;">
                <h1>❌ Authentication Failed</h1>
                <p>Error: ${error}</p>
                <p>You can close this window and try again.</p>
              </body>
            </html>
          `);
          server.close();
          reject(new Error(`Authentication failed: ${error}`));
          return;
        }
        
        if (authCode) {
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(`
            <html>
              <body style="font-family: Arial; padding: 50px; text-align: center;">
                <h1>✅ Authentication Successful!</h1>
                <p><strong>Token saved successfully!</strong></p>
                <p>You can now close this window and start Cursor.</p>
                <script>setTimeout(() => window.close(), 5000);</script>
              </body>
            </html>
          `);
          server.close();
          resolve(authCode);
          return;
        }
      }
      
      // Handle other requests
      res.writeHead(404);
      res.end('Not found');
    });
    
    server.listen(PORT, 'localhost', () => {
      console.log(`🌐 Local auth server started at http://localhost:${PORT}`);
    });
    
    server.on('error', (err) => {
      reject(new Error(`Failed to start local server: ${err.message}`));
    });
  });
}

// Main authentication function
async function authenticate() {
  try {
    console.log('🚀 Starting authentication process...');
    
    // Build authorization URL
    const authCodeUrlParameters = {
      scopes: graphScopes,
      redirectUri: REDIRECT_URI,
      state: crypto.randomBytes(32).toString('hex'),
    };

    const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);
    
    // Start local server to receive callback
    console.log('🌐 Starting local web server for OAuth callback...');
    const authCodePromise = startAuthServer();
    
    // Open browser
    console.log('🔍 Opening browser for Microsoft sign-in...');
    console.log(`URL: ${authUrl}`);
    
    try {
      if (process.platform === 'win32') {
        // On Windows, we need to properly escape the URL
        // Use start with quotes to handle special characters properly
        spawn('cmd', ['/c', 'start', '""', `"${authUrl}"`], { stdio: 'ignore', detached: true }).unref();
      } else if (process.platform === 'darwin') {
        spawn('open', [authUrl], { stdio: 'ignore', detached: true }).unref();
      } else {
        spawn('xdg-open', [authUrl], { stdio: 'ignore', detached: true }).unref();
      }
      console.log('🌐 Browser opened successfully');
    } catch (error) {
      console.log(`⚠️  Could not open browser automatically. Please visit: ${authUrl}`);
    }
    
    console.log('⏳ Waiting for you to complete sign-in in your browser...');
    
    // Wait for auth code from callback
    const authCode = await authCodePromise;
    
    // Exchange authorization code for token
    console.log('🔄 Exchanging authorization code for access token...');
    const tokenRequest = {
      code: authCode,
      scopes: graphScopes,
      redirectUri: REDIRECT_URI,
    };

    const authResult = await cca.acquireTokenByCode(tokenRequest);
    
    if (!authResult) {
      throw new Error('Failed to acquire token via authorization code flow');
    }
    
    // Save token to file
    const tokenData = {
      accessToken: authResult.accessToken,
      expiresOn: authResult.expiresOn,
      account: authResult.account,
      scopes: authResult.scopes,
      tokenType: authResult.tokenType,
      refreshToken: authResult.refreshToken,
      timestamp: Date.now()
    };
    
    fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokenData, null, 2));
    
    console.log('\\n✅ Authentication successful!');
    console.log(`👤 Signed in as: ${authResult.account?.name || authResult.account?.username}`);
    console.log(`🎯 Token saved to: ${TOKEN_FILE}`);
    console.log(`⏰ Token expires: ${authResult.expiresOn}`);
    console.log('\\n🚀 You can now start Cursor and use OneNote MCP commands!');
    
    return authResult;
  } catch (error) {
    console.error(`❌ Authentication failed: ${error.message}`);
    process.exit(1);
  }
}

// Check if we already have a valid token
function checkExistingToken() {
  try {
    if (fs.existsSync(TOKEN_FILE)) {
      const tokenData = JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8'));
      const expiresOn = new Date(tokenData.expiresOn);
      const now = new Date();
      
      if (expiresOn > now) {
        console.log('✅ Valid token found!');
        console.log(`👤 Account: ${tokenData.account?.name || tokenData.account?.username}`);
        console.log(`⏰ Expires: ${expiresOn}`);
        console.log('🚀 You can start Cursor and use OneNote MCP commands!');
        return true;
      } else {
        console.log('⏰ Existing token has expired, re-authentication needed');
        return false;
      }
    }
    return false;
  } catch (error) {
    console.log('🔄 No valid token found, authentication needed');
    return false;
  }
}

// Main execution
async function main() {
  console.log('🔐 OneNote MCP Authentication Tool');
  console.log('=================================');
  
  // Check for existing valid token
  if (checkExistingToken()) {
    return;
  }
  
  // Perform authentication
  await authenticate();
}

main().catch((error) => {
  console.error('❌ Fatal error:', error);
  process.exit(1);
});
