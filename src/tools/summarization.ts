import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { log } from '../utils/logger';

export function registerSummarizationTools(mcpServer: McpServer, graphClient: Client): void {
  log('Registering summarization tools...');

  mcpServer.registerTool('summarize_page', {
    description: 'Summarize the content of a specific OneNote page',
    inputSchema: {
      pageId: z.string().describe('The ID of the page to summarize'),
      maxLength: z.number().optional().describe('Maximum length of summary (default: 300 words)'),
    },
  }, async ({ pageId, maxLength = 300 }) => {
    log(`=== summarize_page called for page: ${pageId} ===`);
    
    try {
      // Get page metadata first
      const pageResponse = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      
      // Get page content
      const contentResponse = await graphClient.api(`/me/onenote/pages/${pageId}/content`).get();
      
      // Handle ReadableStream response from Graph API
      let htmlContent: string;
      if (contentResponse && contentResponse.constructor.name === 'ReadableStream') {
        // Convert ReadableStream to string
        const chunks = [];
        const reader = contentResponse.getReader();
        
        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            chunks.push(value);
          }
          
          // Combine chunks and decode
          const combined = new Uint8Array(chunks.reduce((acc: number, chunk: Uint8Array) => acc + chunk.length, 0));
          let offset = 0;
          for (const chunk of chunks) {
            combined.set(chunk, offset);
            offset += chunk.length;
          }
          
          htmlContent = new TextDecoder().decode(combined);
        } finally {
          reader.releaseLock();
        }
      } else if (typeof contentResponse === 'string') {
        htmlContent = contentResponse;
      } else {
        htmlContent = JSON.stringify(contentResponse);
      }
      
      // Extract text from HTML content
      const textContent = htmlContent
        .replace(/<style[^>]*>.*?<\/style>/gis, '') // Remove style tags
        .replace(/<script[^>]*>.*?<\/script>/gis, '') // Remove script tags
        .replace(/<[^>]*>/g, ' ') // Remove all HTML tags
        .replace(/\s+/g, ' ') // Normalize whitespace
        .trim();
      
      // Basic text summarization (extractive approach)
      const sentences = textContent.split(/[.!?]+/).filter((s: string) => s.trim().length > 20);
      const wordCount = textContent.split(/\s+/).length;
      
      let summary;
      if (wordCount <= maxLength) {
        summary = textContent;
      } else {
        // Take first few sentences that fit within maxLength
        let currentLength = 0;
        const selectedSentences = [];
        
        for (const sentence of sentences) {
          const sentenceWords = sentence.trim().split(/\s+/).length;
          if (currentLength + sentenceWords <= maxLength) {
            selectedSentences.push(sentence.trim());
            currentLength += sentenceWords;
          } else {
            break;
          }
        }
        
        summary = selectedSentences.join('. ') + (selectedSentences.length < sentences.length ? '...' : '');
      }
      
      log(`Page summarized: ${wordCount} words -> ${summary.split(/\s+/).length} words`);
      
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              pageTitle: pageResponse.title,
              pageId: pageId,
              createdDate: pageResponse.createdDateTime,
              lastModified: pageResponse.lastModifiedDateTime,
              originalWordCount: wordCount,
              summaryWordCount: summary.split(/\s+/).length,
              summary: summary,
              webUrl: pageResponse.links?.oneNoteWebUrl?.href
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      log(`ERROR in summarize_page: ${error instanceof Error ? error.message : 'Unknown error'}`);
      throw new Error(`Failed to summarize page: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  });

  mcpServer.registerTool('summarize_section', {
    description: 'Summarize all pages in a OneNote section',
    inputSchema: {
      sectionId: z.string().describe('The ID of the section to summarize'),
      maxPagesPerSection: z.number().optional().describe('Maximum number of pages to include (default: 10)'),
    },
  }, async ({ sectionId, maxPagesPerSection = 10 }) => {
    log(`=== summarize_section called for section: ${sectionId} ===`);
    
    try {
      // Get section metadata
      const sectionResponse = await graphClient.api(`/me/onenote/sections/${sectionId}`).get();
      
      // Get pages in section
      const pagesResponse = await graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get();
      const pages = pagesResponse.value.slice(0, maxPagesPerSection);
      
      const pageSummaries = [];
      
      for (const page of pages) {
        try {
          // Get page content
          const contentResponse = await graphClient.api(`/me/onenote/pages/${page.id}/content`).get();
          
          // Handle ReadableStream response from Graph API
          let htmlContent: string;
          if (contentResponse && contentResponse.constructor.name === 'ReadableStream') {
            // Convert ReadableStream to string
            const chunks = [];
            const reader = contentResponse.getReader();
            
            try {
              while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                chunks.push(value);
              }
              
              const combined = new Uint8Array(chunks.reduce((acc: number, chunk: Uint8Array) => acc + chunk.length, 0));
              let offset = 0;
              for (const chunk of chunks) {
                combined.set(chunk, offset);
                offset += chunk.length;
              }
              
              htmlContent = new TextDecoder().decode(combined);
            } finally {
              reader.releaseLock();
            }
          } else if (typeof contentResponse === 'string') {
            htmlContent = contentResponse;
          } else {
            htmlContent = 'No content available';
          }
          
          // Extract and summarize text
          const textContent = htmlContent
            .replace(/<style[^>]*>.*?<\/style>/gis, '')
            .replace(/<script[^>]*>.*?<\/script>/gis, '')
            .replace(/<[^>]*>/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
          
          const sentences = textContent.split(/[.!?]+/).filter((s: string) => s.trim().length > 10);
          const firstSentence = sentences[0]?.trim() || 'No content available';
          
          pageSummaries.push({
            title: page.title,
            createdDate: page.createdDateTime,
            lastModified: page.lastModifiedDateTime,
            summary: firstSentence.length > 200 ? firstSentence.substring(0, 200) + '...' : firstSentence,
            webUrl: page.links?.oneNoteWebUrl?.href
          });
          
        } catch (pageError) {
          log(`Could not process page: ${page.title}`);
          pageSummaries.push({
            title: page.title,
            createdDate: page.createdDateTime,
            lastModified: page.lastModifiedDateTime,
            summary: 'Content could not be accessed',
            webUrl: page.links?.oneNoteWebUrl?.href
          });
        }
      }
      
      log(`Section summarized: ${pages.length} pages processed`);
      
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              sectionName: sectionResponse.displayName,
              sectionId: sectionId,
              totalPages: pagesResponse.value.length,
              summarizedPages: pageSummaries.length,
              pageSummaries: pageSummaries
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      log(`ERROR in summarize_section: ${error instanceof Error ? error.message : 'Unknown error'}`);
      throw new Error(`Failed to summarize section: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  });

  log('Summarization tools registered successfully');
}
