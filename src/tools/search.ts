import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { log } from '../utils/logger';

export function registerSearchTools(mcpServer: McpServer, graphClient: Client): void {
  log('Registering search functionality...');

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

  log('Search tools registered successfully');
}

