#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { ConfidentialClientApplication, AuthenticationResult } from '@azure/msal-node';
import { z } from 'zod';
import { spawn } from 'child_process';
import * as http from 'http';
import * as url from 'url';
import * as crypto from 'crypto';

// Create MCP server
const mcpServer = new McpServer({
  name: 'mcp-server-onenote-delegated',
  version: '0.1.1',
});

// Initialize Azure app configuration
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;

if (!tenantId || !clientId) {
  throw new Error('AZURE_TENANT_ID and AZURE_CLIENT_ID must be provided via environment variables');
}

console.log('🔧 OneNote MCP Server (Delegated Auth) starting...');
console.log('📝 This version uses delegated permissions with user sign-in');

// Configuration for local web server auth
const PORT = 3001;
const REDIRECT_URI = `http://localhost:${PORT}/auth/callback`;

// MSAL configuration for authorization code flow
const msalConfig = {
  auth: {
    clientId: clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET || 'not-needed-for-public-client',
  },
};

const cca = new ConfidentialClientApplication(msalConfig);

// Token storage
let cachedTokenResponse: AuthenticationResult | null = null;

// Microsoft Graph scopes for delegated permissions
const graphScopes = [
  'https://graph.microsoft.com/Notes.ReadWrite',
  'https://graph.microsoft.com/User.Read',
];

// Helper function to start local web server for OAuth callback
function startAuthServer(): Promise<string> {
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const parsedUrl = url.parse(req.url || '', true);
      
      if (parsedUrl.pathname === '/auth/callback') {
        const authCode = parsedUrl.query.code as string;
        const error = parsedUrl.query.error as string;
        
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
                <p>You can close this window and return to your application.</p>
                <script>setTimeout(() => window.close(), 3000);</script>
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

// Authentication helper functions
async function getAccessToken(): Promise<string> {
  try {
    // Try to get token silently first (from cache)
    if (cachedTokenResponse?.account) {
      const silentRequest = {
        account: cachedTokenResponse.account,
        scopes: graphScopes,
      };
      
      try {
        const silentResult = await cca.acquireTokenSilent(silentRequest);
        console.log('🔄 Using cached authentication token');
        return silentResult.accessToken;
      } catch (silentError) {
        console.log('🔄 Cached token expired, requesting new authentication...');
      }
    }

    // Use authorization code flow with local web server
    console.log('\n🔐 USER AUTHENTICATION REQUIRED');
    console.log('================================');
    console.log('Opening browser for Microsoft sign-in...');
    console.log('================================');
    
    // Generate state for security
    const state = crypto.randomBytes(32).toString('hex');
    
    // Build authorization URL
    const authCodeUrlParameters = {
      scopes: graphScopes,
      redirectUri: REDIRECT_URI,
      state: state,
    };

    const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);
    
    // Start local server to receive callback
    const authCodePromise = startAuthServer();
    
    // Open browser
    console.log(`🌐 Opening browser to: ${authUrl}`);
    const opener = process.platform === 'win32' ? 'start' : process.platform === 'darwin' ? 'open' : 'xdg-open';
    try {
      spawn(opener, [authUrl], { stdio: 'ignore', detached: true }).unref();
    } catch (error) {
      console.log(`⚠️  Could not open browser automatically. Please visit: ${authUrl}`);
    }
    
    console.log('⏳ Waiting for you to complete sign-in in your browser...\n');
    
    // Wait for auth code from callback
    const authCode = await authCodePromise;
    
    // Exchange authorization code for token
    const tokenRequest = {
      code: authCode,
      scopes: graphScopes,
      redirectUri: REDIRECT_URI,
    };

    const authResult = await cca.acquireTokenByCode(tokenRequest);
    
    if (!authResult) {
      throw new Error('Failed to acquire token via authorization code flow');
    }
    
    cachedTokenResponse = authResult;
    
    console.log('✅ Authentication successful!');
    console.log(`👤 Signed in as: ${authResult.account?.name || authResult.account?.username}`);
    console.log('🎯 OneNote MCP Server is ready!\n');
    
    return authResult.accessToken;
  } catch (error) {
    throw new Error(`Authentication failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

// Initialize Microsoft Graph client with delegated auth
const graphClient = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      return await getAccessToken();
    }
  }
});

// Register notebook management tools
mcpServer.registerTool('list_notebooks', {
  description: 'List all OneNote notebooks for the authenticated user',
  inputSchema: {},
}, async () => {
  try {
    const response = await graphClient.api('/me/onenote/notebooks').get();
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response.value, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list notebooks: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

mcpServer.registerTool('get_notebook', {
  description: 'Get details of a specific OneNote notebook',
  inputSchema: {
    id: z.string().describe('The ID of the notebook'),
  },
}, async ({ id }) => {
  try {
    const response = await graphClient.api(`/me/onenote/notebooks/${id}`).get();
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to get notebook ${id}: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

mcpServer.registerTool('create_notebook', {
  description: 'Create a new OneNote notebook for the authenticated user',
  inputSchema: {
    name: z.string().describe('The name of the new notebook'),
  },
}, async ({ name }) => {
  try {
    const response = await graphClient.api('/me/onenote/notebooks').post({
      displayName: name,
    });
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create notebook: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

// Register section management tools
mcpServer.registerTool('list_sections', {
  description: 'List sections in a OneNote notebook',
  inputSchema: {
    notebookId: z.string().describe('The ID of the notebook'),
  },
}, async ({ notebookId }) => {
  try {
    const response = await graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`).get();
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response.value, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list sections: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

mcpServer.registerTool('create_section', {
  description: 'Create a new section in a OneNote notebook',
  inputSchema: {
    notebookId: z.string().describe('The ID of the notebook'),
    name: z.string().describe('The name of the new section'),
  },
}, async ({ notebookId, name }) => {
  try {
    const response = await graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`).post({
      displayName: name,
    });
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create section: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

// Register page management tools
mcpServer.registerTool('list_pages', {
  description: 'List pages in a OneNote section',
  inputSchema: {
    sectionId: z.string().describe('The ID of the section'),
  },
}, async ({ sectionId }) => {
  try {
    const response = await graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get();
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response.value, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list pages: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

mcpServer.registerTool('get_page', {
  description: 'Get details of a specific OneNote page',
  inputSchema: {
    id: z.string().describe('The ID of the page'),
  },
}, async ({ id }) => {
  try {
    const response = await graphClient.api(`/me/onenote/pages/${id}`).get();
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to get page ${id}: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

mcpServer.registerTool('create_page', {
  description: 'Create a new page in a OneNote section',
  inputSchema: {
    sectionId: z.string().describe('The ID of the section'),
    title: z.string().describe('The title of the new page'),
    content: z.string().optional().describe('HTML content for the page'),
  },
}, async ({ sectionId, title, content = '' }) => {
  try {
    const html = `
<!DOCTYPE html>
<html>
  <head>
    <title>${title}</title>
  </head>
  <body>
    <h1>${title}</h1>
    ${content}
  </body>
</html>`;
    
    const response = await graphClient.api(`/me/onenote/sections/${sectionId}/pages`)
      .header('Content-Type', 'text/html')
      .post(html);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create page: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

// Start server if run directly
async function main() {
  const transport = new StdioServerTransport();
  await mcpServer.connect(transport);
  console.log('🚀 OneNote MCP Server (Delegated Auth) is running...');
}

if (require.main === module) {
  main().catch((error) => {
    console.error('❌ Server error:', error);
    process.exit(1);
  });
}