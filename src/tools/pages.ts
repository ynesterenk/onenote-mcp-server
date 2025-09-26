import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { log } from '../utils/logger';

export function registerPageTools(mcpServer: McpServer, graphClient: Client): void {
  log('Registering page management tools...');

  mcpServer.registerTool('list_pages', {
    description: 'List pages in a OneNote section',
    inputSchema: {
      sectionId: z.string().describe('The ID of the section'),
    },
  }, async ({ sectionId }) => {
    log(`=== list_pages tool called for section: ${sectionId} ===`);
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
    log(`=== get_page tool called for ID: ${id} ===`);
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
    log(`=== create_page tool called for section: ${sectionId}, title: ${title} ===`);
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

  log('Page tools registered successfully');
}

