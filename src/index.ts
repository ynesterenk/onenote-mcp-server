#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';

// Create MCP server
const mcpServer = new McpServer({
  name: 'mcp-server-onenote',
  version: '0.1.1',
});

// Initialize Azure credentials
const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

if (!tenantId || !clientId || !clientSecret) {
  throw new Error('Azure credentials must be provided via environment variables');
}

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

// Initialize Microsoft Graph client
const graphClient = Client.initWithMiddleware({
  authProvider: {
    async getAccessToken() {
      const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
      return tokenResponse?.token || '';
    }
  }
});

// Register notebook management tools
mcpServer.registerTool('list_notebooks', {
  description: 'List OneNote notebooks for a specific user',
  inputSchema: {
    userId: z.string().describe('User ID or email address (required for client credentials flow)'),
  },
}, async ({ userId }) => {
  try {
    const apiPath = `/users/${userId}/onenote/notebooks`;
    const response = await graphClient.api(apiPath).get();
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
    userId: z.string().optional().describe('User ID or email address'),
  },
}, async ({ id, userId }) => {
  try {
    const apiPath = userId ? `/users/${userId}/onenote/notebooks/${id}` : `/me/onenote/notebooks/${id}`;
    const response = await graphClient.api(apiPath).get();
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
  description: 'Create a new OneNote notebook',
  inputSchema: {
    name: z.string().describe('The name of the new notebook'),
    userId: z.string().optional().describe('User ID or email address'),
  },
}, async ({ name, userId }) => {
  try {
    const apiPath = userId ? `/users/${userId}/onenote/notebooks` : `/me/onenote/notebooks`;
    const response = await graphClient.api(apiPath).post({
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
    userId: z.string().optional().describe('User ID or email address'),
  },
}, async ({ notebookId, userId }) => {
  try {
    const apiPath = userId ? `/users/${userId}/onenote/notebooks/${notebookId}/sections` : `/me/onenote/notebooks/${notebookId}/sections`;
    const response = await graphClient.api(apiPath).get();
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
    userId: z.string().optional().describe('User ID or email address'),
  },
}, async ({ notebookId, name, userId }) => {
  try {
    const apiPath = userId ? `/users/${userId}/onenote/notebooks/${notebookId}/sections` : `/me/onenote/notebooks/${notebookId}/sections`;
    const response = await graphClient.api(apiPath).post({
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
    userId: z.string().optional().describe('User ID or email address'),
  },
}, async ({ sectionId, userId }) => {
  try {
    const apiPath = userId ? `/users/${userId}/onenote/sections/${sectionId}/pages` : `/me/onenote/sections/${sectionId}/pages`;
    const response = await graphClient.api(apiPath).get();
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
    userId: z.string().optional().describe('User ID or email address'),
  },
}, async ({ id, userId }) => {
  try {
    const apiPath = userId ? `/users/${userId}/onenote/pages/${id}` : `/me/onenote/pages/${id}`;
    const response = await graphClient.api(apiPath).get();
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
    userId: z.string().optional().describe('User ID or email address'),
  },
}, async ({ sectionId, title, content = '', userId }) => {
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
    
    const apiPath = userId ? `/users/${userId}/onenote/sections/${sectionId}/pages` : `/me/onenote/sections/${sectionId}/pages`;
    const response = await graphClient.api(apiPath)
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
  console.log('OneNote MCP server is running...');
}

if (require.main === module) {
  main().catch((error) => {
    console.error('Server error:', error);
    process.exit(1);
  });
}