#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { TokenManager } from './auth/tokenManager';
import { createGraphClient } from './utils/graphClient';
import { registerNotebookTools } from './tools/notebooks';
import { registerSectionTools } from './tools/sections';
import { registerPageTools } from './tools/pages';
import { registerSearchTools } from './tools/search';
import { log } from './utils/logger';

// Start logging
log('=== OneNote MCP Server (Delegated Auth) Starting ===');
log(`Process ID: ${process.pid}`);
log(`Working directory: ${process.cwd()}`);

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

// Initialize authentication and Graph client
const tokenManager = new TokenManager(tenantId, clientId);
const graphClient = createGraphClient(tokenManager);

// Register all tools
log('Registering all MCP tools...');
registerNotebookTools(mcpServer, graphClient);
registerSectionTools(mcpServer, graphClient);
registerPageTools(mcpServer, graphClient);
registerSearchTools(mcpServer, graphClient);
log('All tools registered successfully');

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