import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { log } from '../utils/logger';

// Helper function to calculate relevance score
function calculateRelevance(query: string, title: string, content?: string): number {
  const queryLower = query.toLowerCase();
  const titleLower = title.toLowerCase();
  const contentLower = content?.toLowerCase() || '';
  
  let score = 0;
  
  // Exact title match gets highest score
  if (titleLower === queryLower) score += 100;
  
  // Title contains exact phrase
  if (titleLower.includes(queryLower)) score += 50;
  
  // Content contains exact phrase
  if (contentLower.includes(queryLower)) score += 30;
  
  // Individual word matches in title
  const queryWords = queryLower.split(/\s+/);
  queryWords.forEach(word => {
    if (titleLower.includes(word)) score += 20;
    if (contentLower.includes(word)) score += 10;
  });
  
  // Semantic/fuzzy matching for common variations
  const semanticMappings = {
    'agile': ['scrum', 'sprint', 'kanban', 'team lead', 'leadership'],
    'ai': ['artificial intelligence', 'machine learning', 'ml', 'neural', 'ai-900'],
    'azure': ['cloud', 'microsoft', 'az-305', 'certification'],
    'migration': ['move', 'transfer', 'upgrade', 'moderniz'],
    'architecture': ['design', 'pattern', 'structure', 'microservices'],
    'team': ['group', 'squad', 'crew', 'leadership', 'management'],
    'lead': ['leader', 'management', 'tlg', 'manager']
  };
  
  Object.entries(semanticMappings).forEach(([key, synonyms]) => {
    if (queryLower.includes(key)) {
      synonyms.forEach(synonym => {
        if (titleLower.includes(synonym)) score += 15;
        if (contentLower.includes(synonym)) score += 8;
      });
    }
  });
  
  return score;
}

// Helper function to extract content preview
function extractContentPreview(content: string, query: string, maxLength: number = 200): string {
  if (!content) return '';
  
  const queryLower = query.toLowerCase();
  const contentLower = content.toLowerCase();
  
  // Find the first occurrence of query in content
  const index = contentLower.indexOf(queryLower);
  if (index === -1) return content.substring(0, maxLength) + '...';
  
  // Extract context around the match
  const start = Math.max(0, index - 50);
  const end = Math.min(content.length, index + queryLower.length + 150);
  
  let preview = content.substring(start, end);
  if (start > 0) preview = '...' + preview;
  if (end < content.length) preview = preview + '...';
  
  return preview;
}

// Generic date period parser
function parsePeriod(period: string): { start: Date; end: Date } | null {
  const periodLower = period.toLowerCase().trim();
  
  // Match patterns like "Q1 2020", "q1 2020", "2020 Q1", "2020q1"
  const quarterYearMatch = periodLower.match(/q(\d)\s*(\d{4})|(\d{4})\s*q(\d)/);
  if (quarterYearMatch) {
    const quarter = parseInt(quarterYearMatch[1] || quarterYearMatch[4]);
    const year = parseInt(quarterYearMatch[2] || quarterYearMatch[3]);
    
    if (quarter >= 1 && quarter <= 4 && year >= 1900 && year <= 2100) {
      const quarterStartMonth = (quarter - 1) * 3; // Q1=0, Q2=3, Q3=6, Q4=9
      const quarterEndMonth = quarterStartMonth + 2; // Q1=2, Q2=5, Q3=8, Q4=11
      
      const startDate = new Date(year, quarterStartMonth, 1);
      const endDate = new Date(year, quarterEndMonth + 1, 0); // Last day of quarter
      
      return { start: startDate, end: endDate };
    }
  }
  
  // Match full year patterns like "2020", "2021"
  const yearMatch = periodLower.match(/^(\d{4})$/);
  if (yearMatch) {
    const year = parseInt(yearMatch[1]);
    if (year >= 1900 && year <= 2100) {
      return {
        start: new Date(year, 0, 1),
        end: new Date(year, 11, 31)
      };
    }
  }
  
  // Match month year patterns like "January 2020", "Jan 2020", "2020-01"
  const monthYearMatch = periodLower.match(/^(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s*(\d{4})$|^(\d{4})-(\d{1,2})$/);
  if (monthYearMatch) {
    let month: number;
    let year: number;
    
    if (monthYearMatch[3] && monthYearMatch[4]) {
      // Format: "2020-01"
      year = parseInt(monthYearMatch[3]);
      month = parseInt(monthYearMatch[4]) - 1; // Convert to 0-based
    } else {
      // Format: "January 2020"
      year = parseInt(monthYearMatch[2]);
      const monthName = monthYearMatch[1];
      const monthMap: { [key: string]: number } = {
        'january': 0, 'jan': 0, 'february': 1, 'feb': 1, 'march': 2, 'mar': 2,
        'april': 3, 'apr': 3, 'may': 4, 'june': 5, 'jun': 5,
        'july': 6, 'jul': 6, 'august': 7, 'aug': 7, 'september': 8, 'sep': 8,
        'october': 9, 'oct': 9, 'november': 10, 'nov': 10, 'december': 11, 'dec': 11
      };
      month = monthMap[monthName];
    }
    
    if (month >= 0 && month <= 11 && year >= 1900 && year <= 2100) {
      return {
        start: new Date(year, month, 1),
        end: new Date(year, month + 1, 0) // Last day of month
      };
    }
  }
  
  // Match relative periods like "last 6 months", "recent"
  if (periodLower.includes('recent') || periodLower.includes('last year')) {
    const oneYearAgo = new Date();
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    return {
      start: oneYearAgo,
      end: new Date()
    };
  }
  
  const lastMonthsMatch = periodLower.match(/last\s*(\d+)\s*months?/);
  if (lastMonthsMatch) {
    const months = parseInt(lastMonthsMatch[1]);
    const startDate = new Date();
    startDate.setMonth(startDate.getMonth() - months);
    return {
      start: startDate,
      end: new Date()
    };
  }
  
  log(`Could not parse period: "${period}"`);
  return null;
}

export function registerSearchTools(mcpServer: McpServer, graphClient: Client): void {
  log('Registering enhanced search functionality...');

  mcpServer.registerTool('search_pages', {
    description: 'AI-powered semantic search across all OneNote notebooks - understands meaning, not just keywords',
    inputSchema: {
      query: z.string().describe('Natural language search query (e.g., "leadership training", "Azure architecture", "client projects")'),
      includeContent: z.boolean().optional().describe('Include page content in search (slower but more accurate)'),
      maxResults: z.number().optional().describe('Maximum number of results to return (default: 10)'),
    },
  }, async ({ query, includeContent = false, maxResults = 10 }) => {
    log(`=== search_pages tool called with query: "${query}" ===`);
    log(`Options: includeContent=${includeContent}, maxResults=${maxResults}`);
    
    try {
      log('Starting AI-powered semantic search...');
      
      // Get all notebooks first
      const notebooksResponse = await graphClient.api('/me/onenote/notebooks').get();
      const searchResults = [];
      
      let totalPagesSearched = 0;
      
      // Search through each notebook
      for (const notebook of notebooksResponse.value) {
        log(`Searching in notebook: ${notebook.displayName}`);
        
        // Get sections for this notebook
        const sectionsResponse = await graphClient.api(`/me/onenote/notebooks/${notebook.id}/sections`).get();
        
        // Search through each section
        for (const section of sectionsResponse.value) {
          log(`Searching in section: ${section.displayName}`);
          
          // Get pages for this section with pagination
          let allPages = [];
          let nextLink = `/me/onenote/sections/${section.id}/pages`;
          
          while (nextLink) {
            const pagesResponse = await graphClient.api(nextLink).get();
            allPages.push(...pagesResponse.value);
            
            nextLink = pagesResponse['@odata.nextLink'] ? 
              pagesResponse['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : 
              null;
          }
          
          // Analyze each page for relevance
          for (const page of allPages) {
            totalPagesSearched++;
            
            let pageContent = '';
            
            // Optionally fetch full page content for deeper search
            if (includeContent) {
              try {
                const contentResponse = await graphClient.api(`/me/onenote/pages/${page.id}/content`).get();
                // Extract text from HTML content (basic text extraction)
                pageContent = contentResponse.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
              } catch (contentError) {
                log(`Could not fetch content for page: ${page.title}`);
              }
            }
            
            // Calculate semantic relevance score
            const relevanceScore = calculateRelevance(query, page.title, pageContent);
            
            // Include results with relevance score > 0
            if (relevanceScore > 0) {
              searchResults.push({
                relevanceScore,
                notebook: notebook.displayName,
                section: section.displayName,
                page: {
                  id: page.id,
                  title: page.title,
                  createdDateTime: page.createdDateTime,
                  lastModifiedDateTime: page.lastModifiedDateTime,
                  webUrl: page.links?.oneNoteWebUrl?.href,
                  clientUrl: page.links?.oneNoteClientUrl?.href,
                  preview: includeContent ? extractContentPreview(pageContent, query) : null
                }
              });
            }
          }
        }
      }
      
      // Sort by relevance score (highest first)
      searchResults.sort((a, b) => b.relevanceScore - a.relevanceScore);
      
      // Limit results
      const limitedResults = searchResults.slice(0, maxResults);
      
      log(`Semantic search completed. Searched ${totalPagesSearched} pages, found ${searchResults.length} relevant results, returning top ${limitedResults.length}`);
      
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              query: query,
              searchType: 'semantic',
              totalPagesSearched,
              totalRelevantResults: searchResults.length,
              returnedResults: limitedResults.length,
              results: limitedResults
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      log(`ERROR in search_pages: ${error instanceof Error ? error.message : 'Unknown error'}`);
      throw new Error(`Failed to search pages: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  });

  // Add related content finder
  mcpServer.registerTool('find_related_content', {
    description: 'Find content related to a specific topic or project across all notebooks',
    inputSchema: {
      topic: z.string().describe('Topic or project name (e.g., "LSEG", "Azure", "leadership")'),
      timeframe: z.string().optional().describe('Time period to focus on (e.g., "2021", "last year", "recent")'),
    },
  }, async ({ topic, timeframe }) => {
    log(`=== find_related_content called for topic: "${topic}" ===`);
    
    try {
      // Use semantic search with expanded terms
      const expandedQuery = topic.toLowerCase();
      const results = [];
      
      // Get all notebooks
      const notebooksResponse = await graphClient.api('/me/onenote/notebooks').get();
      
      for (const notebook of notebooksResponse.value) {
        const sectionsResponse = await graphClient.api(`/me/onenote/notebooks/${notebook.id}/sections`).get();
        
        for (const section of sectionsResponse.value) {
          // Check if section name is related to topic
          const sectionRelevance = calculateRelevance(topic, section.displayName);
          
          if (sectionRelevance > 0) {
            // Handle pagination for pages
            let allPages = [];
            let nextLink = `/me/onenote/sections/${section.id}/pages`;
            
            while (nextLink) {
              const pagesResponse = await graphClient.api(nextLink).get();
              allPages.push(...pagesResponse.value);
              
              nextLink = pagesResponse['@odata.nextLink'] ? 
                pagesResponse['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : 
                null;
            }
            
            for (const page of allPages) {
              // Apply timeframe filter if specified
              if (timeframe) {
                const pageYear = new Date(page.createdDateTime).getFullYear().toString();
                if (timeframe.includes(pageYear) || 
                    (timeframe.includes('recent') && new Date(page.lastModifiedDateTime) > new Date(Date.now() - 365 * 24 * 60 * 60 * 1000))) {
                  // Include if timeframe matches
                } else {
                  continue; // Skip if timeframe doesn't match
                }
              }
              
              const pageRelevance = calculateRelevance(topic, page.title);
              const totalRelevance = sectionRelevance + pageRelevance;
              
              if (totalRelevance > 15) { // Higher threshold for related content
                results.push({
                  relevanceScore: totalRelevance,
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
      }
      
      // Sort by relevance and limit results
      results.sort((a, b) => b.relevanceScore - a.relevanceScore);
      const limitedResults = results.slice(0, 15);
      
      log(`Related content search completed. Found ${results.length} related items`);
      
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              topic: topic,
              timeframe: timeframe || 'all time',
              searchType: 'related_content',
              totalResults: results.length,
              results: limitedResults
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to find related content: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  });

  // Add timeline search
  mcpServer.registerTool('search_by_timeline', {
    description: 'Search and organize content by time periods - great for project retrospectives',
    inputSchema: {
      period: z.string().describe('Time period (e.g., "2021", "last 6 months", "Q3 2021")'),
      topic: z.string().optional().describe('Optional topic filter'),
    },
  }, async ({ period, topic }) => {
    log(`=== search_by_timeline called for period: "${period}" ===`);
    
    try {
      const results = [];
      const notebooksResponse = await graphClient.api('/me/onenote/notebooks').get();
      
      // Parse time period with generic approach
      const parsedPeriod = parsePeriod(period);
      if (!parsedPeriod) {
        log(`ERROR: Could not parse period "${period}"`);
        throw new Error(`Invalid period format: "${period}"`);
      }
      
      const startDate = parsedPeriod.start;
      const endDate = parsedPeriod.end;
      
      log(`Searching for pages between: ${startDate.toISOString()} and ${endDate.toISOString()}`);
      
      let totalPagesChecked = 0;
      let pagesInRange = 0;
      
      // Sequential API calls to respect rate limits
      for (const notebook of notebooksResponse.value) {
        const sectionsResponse = await graphClient.api(`/me/onenote/notebooks/${notebook.id}/sections`).get();
        
        for (const section of sectionsResponse.value) {
          // Handle pagination for pages
          let allPages = [];
          let nextLink = `/me/onenote/sections/${section.id}/pages`;
          
          while (nextLink) {
            const pagesResponse = await graphClient.api(nextLink).get();
            allPages.push(...pagesResponse.value);
            
            nextLink = pagesResponse['@odata.nextLink'] ? 
              pagesResponse['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : 
              null;
          }
          
          for (const page of allPages) {
            totalPagesChecked++;
            const pageDate = new Date(page.createdDateTime);
            
            // Debug logging for date comparison
            if (totalPagesChecked <= 5) { // Only log first 5 for debugging
              log(`Page "${page.title}" created: ${pageDate.toISOString()}`);
            }
            
            // Check if page falls within time period (fix: remove the faulty logic)
            const inTimeRange = pageDate >= startDate && pageDate <= endDate;
            
            if (inTimeRange) {
              pagesInRange++;
              log(`✓ Page "${page.title}" is IN range (${pageDate.toDateString()})`);
              
              // Apply topic filter if specified
              let topicMatch = true;
              if (topic) {
                topicMatch = calculateRelevance(topic, page.title) > 0;
              }
              
              if (topicMatch) {
                results.push({
                  notebook: notebook.displayName,
                  section: section.displayName,
                  date: page.createdDateTime,
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
            } else if (totalPagesChecked <= 5) { // Debug first 5 pages
              log(`✗ Page "${page.title}" is OUT of range (${pageDate.toDateString()})`);
            }
          }
        }
      }
      
      log(`Checked ${totalPagesChecked} total pages, found ${pagesInRange} in date range, returning ${results.length} matching results`);
      
      // Sort by date (newest first)
      results.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
      
      log(`Timeline search completed. Found ${results.length} pages in time period`);
      
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              period: period,
              topic: topic || 'all topics',
              searchType: 'timeline',
              totalResults: results.length,
              results: results
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to search by timeline: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  });

  log('Enhanced search tools registered successfully');
}

