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
import * as fs from 'fs';
import * as path from 'path';

// Setup logging
const LOG_FILE = path.join(__dirname, '..', 'mcp-server.log');

function log(message: string) {
  const timestamp = new Date().toISOString();
  const logMessage = `${timestamp}: ${message}\n`;
  
  // Append to log file
  try {
    fs.appendFileSync(LOG_FILE, logMessage);
  } catch (error) {
    // Ignore logging errors
  }
  
  // Also log to console if not in Cursor's STDIO mode
  if (process.env.NODE_ENV !== 'production') {
    console.log(message);
  }
}

// Start logging
log('=== OneNote MCP Server (Delegated Auth) Starting ===');
log(`Process ID: ${process.pid}`);
log(`Working directory: ${process.cwd()}`);
log(`Log file: ${LOG_FILE}`);

// Create MCP server
log('Creating MCP server...');
const mcpServer = new McpServer({
  name: 'mcp-server-onenote-delegated',
  version: '0.1.1',
});
log('MCP server created successfully');

// Initialize Azure app configuration
log('Reading environment variables...');
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;

log(`AZURE_TENANT_ID: ${tenantId ? 'SET' : 'NOT SET'}`);
log(`AZURE_CLIENT_ID: ${clientId ? 'SET' : 'NOT SET'}`);

if (!tenantId || !clientId) {
  log('ERROR: Missing required environment variables');
  throw new Error('AZURE_TENANT_ID and AZURE_CLIENT_ID must be provided via environment variables');
}

log('Environment variables validated successfully');
console.log('🔧 OneNote MCP Server (Delegated Auth) starting...');
console.log('📝 This version uses delegated permissions with user sign-in');

// Configuration for local web server auth
const PORT = 3001;
const REDIRECT_URI = `http://localhost:${PORT}/auth/callback`;

// MSAL configuration for authorization code flow
log('Configuring MSAL...');
const msalConfig = {
  auth: {
    clientId: clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET || 'not-needed-for-public-client',
  },
};

log(`MSAL Authority: ${msalConfig.auth.authority}`);
log('Creating ConfidentialClientApplication...');

let cca: ConfidentialClientApplication;
try {
  cca = new ConfidentialClientApplication(msalConfig);
  log('MSAL client created successfully');
} catch (error) {
  log(`ERROR creating MSAL client: ${error instanceof Error ? error.message : 'Unknown error'}`);
  throw error;
}

// Token storage
let cachedTokenResponse: AuthenticationResult | null = null;

// Microsoft Graph scopes for delegated permissions
const graphScopes = [
  'Notes.ReadWrite',
  'User.Read',
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
  log('=== getAccessToken called ===');
  
  try {
    // Try to read pre-saved token from file
    const TOKEN_FILE = path.join(__dirname, '..', '.auth-token');
    log(`Looking for token file: ${TOKEN_FILE}`);
    
    if (fs.existsSync(TOKEN_FILE)) {
      log('Token file found, reading...');
      const tokenData = JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8'));
      
      const expiresOn = new Date(tokenData.expiresOn);
      const now = new Date();
      
      log(`Token expires: ${expiresOn}`);
      log(`Current time: ${now}`);
      
      if (expiresOn > now) {
        log('Using valid pre-saved token');
        return tokenData.accessToken;
      } else {
        log('Pre-saved token has expired');
        throw new Error('Authentication token has expired. Please run: node auth-tool.js');
      }
    } else {
      log('No token file found');
      throw new Error('No authentication token found. Please run: node auth-tool.js');
    }
  } catch (error) {
    log(`ERROR in getAccessToken: ${error instanceof Error ? error.message : 'Unknown error'}`);
    throw new Error(`Authentication required: ${error instanceof Error ? error.message : 'Unknown error'}. Please run: node auth-tool.js`);
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
  log('=== list_notebooks tool called ===');
  try {
    log('Making Graph API call to /me/onenote/notebooks...');
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

// Register search functionality
mcpServer.registerTool('search_pages', {
  description: 'Search for pages across all OneNote notebooks containing specific text',
  inputSchema: {
    query: z.string().describe('Search query text to find in page titles and content'),
  },
}, async ({ query }) => {
  log(`=== search_pages tool called with query: ${query} ===`);
  try {
    log('Searching across all notebooks and pages...');
    
    // Get all notebooks first
    const notebooksResponse = await graphClient.api('/me/onenote/notebooks').get();
    const allResults = [];
    
    // Search through each notebook
    for (const notebook of notebooksResponse.value) {
      log(`Searching in notebook: ${notebook.displayName}`);
      
      // Get sections for this notebook
      const sectionsResponse = await graphClient.api(`/me/onenote/notebooks/${notebook.id}/sections`).get();
      
      // Search through each section
      for (const section of sectionsResponse.value) {
        log(`Searching in section: ${section.displayName}`);
        
        // Get pages for this section
        const pagesResponse = await graphClient.api(`/me/onenote/sections/${section.id}/pages`).get();
        
        // Filter pages that match the query
        for (const page of pagesResponse.value) {
          const titleMatch = page.title && page.title.toLowerCase().includes(query.toLowerCase());
          
          if (titleMatch) {
            allResults.push({
              notebook: notebook.displayName,
              section: section.displayName,
              page: {
                id: page.id,
                title: page.title,
                createdDateTime: page.createdDateTime,
                lastModifiedDateTime: page.lastModifiedDateTime,
                webUrl: page.links?.oneNoteWebUrl?.href,
                clientUrl: page.links?.oneNoteClientUrl?.href
              }
            });
          }
        }
      }
    }
    
    log(`Search completed. Found ${allResults.length} matching pages`);
    
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            query: query,
            totalResults: allResults.length,
            results: allResults
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    log(`ERROR in search_pages: ${error instanceof Error ? error.message : 'Unknown error'}`);
    throw new Error(`Failed to search pages: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
});

// Start server if run directly
async function main() {
  log('Starting main server function...');
  
  try {
    log('Creating STDIO transport...');
    const transport = new StdioServerTransport();
    log('Transport created, connecting to MCP server...');
    
    await mcpServer.connect(transport);
    log('MCP server connected successfully');
    
    console.log('🚀 OneNote MCP Server (Delegated Auth) is running...');
    log('Server is running and ready for requests');
    
  } catch (error) {
    log(`ERROR in main function: ${error instanceof Error ? error.message : 'Unknown error'}`);
    log(`Stack trace: ${error instanceof Error ? error.stack : 'No stack trace'}`);
    throw error;
  }
}

if (require.main === module) {
  log('Starting server as main module...');
  main().catch((error) => {
    log(`FATAL ERROR: ${error instanceof Error ? error.message : 'Unknown error'}`);
    log(`Stack trace: ${error instanceof Error ? error.stack : 'No stack trace'}`);
    console.error('❌ Server error:', error);
    process.exit(1);
  });
} else {
  log('Module loaded as dependency, not starting server');
}