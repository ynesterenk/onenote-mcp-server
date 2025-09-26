import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { log } from '../utils/logger';

export function registerSectionTools(mcpServer: McpServer, graphClient: Client): void {
  log('Registering section management tools...');

  mcpServer.registerTool('list_sections', {
    description: 'List sections in a OneNote notebook',
    inputSchema: {
      notebookId: z.string().describe('The ID of the notebook'),
    },
  }, async ({ notebookId }) => {
    log(`=== list_sections tool called for notebook: ${notebookId} ===`);
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
    log(`=== create_section tool called for notebook: ${notebookId}, name: ${name} ===`);
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

  log('Section tools registered successfully');
}

