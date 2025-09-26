import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { log } from '../utils/logger';

export function registerNotebookTools(mcpServer: McpServer, graphClient: Client): void {
  log('Registering notebook management tools...');

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
    log(`=== get_notebook tool called for ID: ${id} ===`);
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
    log(`=== create_notebook tool called with name: ${name} ===`);
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

  log('Notebook tools registered successfully');
}

