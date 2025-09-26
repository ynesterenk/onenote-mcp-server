
# OneNote MCP Server
[![smithery badge](https://smithery.ai/badge/@modelcontextprotocol/server-onenote)](https://smithery.ai/server/@modelcontextprotocol/server-onenote)
[![npm version](https://img.shields.io/npm/v/mcp-server-onenote.svg)](https://www.npmjs.com/package/mcp-server-onenote)
[![Verified on MseeP](https://mseep.ai/badge.svg)](https://mseep.ai/app/94d05742-a658-483b-b3e3-f8a3e5ec6e23)

A Model Context Protocol (MCP) server implementation for Microsoft OneNote, enabling AI language models to interact with OneNote through a standardized interface.

## Features

### Core OneNote Management
- **Notebook Management**: List, create, get details, and delete notebooks
- **Section Management**: List, create, get details, and delete sections  
- **Page Management**: List, create, read, update, and delete pages with HTML content

### Advanced Search & AI Features
- **Semantic Search**: AI-powered search that understands meaning, not just keywords
- **Timeline Search**: Search and organize content by time periods (quarters, years, months)
- **Related Content**: Find content related to specific topics or projects
- **Page Summarization**: AI-powered summaries of individual pages or entire sections

## Available Tools

### 🔍 **Search Tools**

#### `search_pages`
**Purpose**: AI-powered semantic search across ALL pages in your notebooks
- **How it works**: Searches every single page regardless of section names
- **Best for**: Finding specific content anywhere ("Java Townhall", "Leading Agile Teams")
- **Coverage**: Complete - searches all pages (566+ pages with pagination)
- **Parameters**: 
  - `query`: Natural language search query
  - `includeContent`: Include page content for deeper search (slower but more accurate)
  - `maxResults`: Maximum results to return (default: 10)

#### `search_by_timeline`
**Purpose**: Search and organize content by time periods - great for project retrospectives
- **How it works**: Filters pages by creation date within specified time periods
- **Best for**: Finding notes from specific time periods ("Q3 2023", "2021", "last 6 months")
- **Date formats supported**: 
  - Quarters: "Q1 2020", "q2 2021", "2023 Q4"
  - Years: "2020", "2021" 
  - Months: "January 2020", "2020-01"
  - Relative: "recent", "last 6 months"
- **Parameters**:
  - `period`: Time period (e.g., "Q3 2023", "2021", "last year")
  - `topic`: Optional topic filter

#### `find_related_content`
**Purpose**: Find content related to specific topics by searching in relevant sections
- **How it works**: Only searches sections whose names match the topic, then searches pages within those sections
- **Best for**: Finding content in topically-related sections ("Google" → searches Google, Google AI, Google - Migration sections)
- **Limitation**: Will miss content if the section name doesn't match the topic
- **Parameters**:
  - `topic`: Topic or project name (e.g., "Google", "Azure", "Microsoft")
  - `timeframe`: Optional time period filter

### 📚 **Notebook Management**

#### `list_notebooks`
**Purpose**: List all your OneNote notebooks
- **Returns**: Array of notebooks with IDs, names, and metadata

#### `get_notebook` 
**Purpose**: Get detailed information about a specific notebook
- **Parameters**: `id` - The notebook ID

#### `create_notebook`
**Purpose**: Create a new OneNote notebook
- **Parameters**: `name` - The name of the new notebook

### 📁 **Section Management**

#### `list_sections`
**Purpose**: List all sections within a specific notebook
- **Parameters**: `notebookId` - The ID of the notebook

#### `create_section`
**Purpose**: Create a new section in a notebook
- **Parameters**: 
  - `notebookId` - The ID of the notebook
  - `name` - The name of the new section

### 📄 **Page Management**

#### `list_pages`
**Purpose**: List ALL pages in a section (with full pagination support)
- **How it works**: Fetches all pages using pagination to handle large sections
- **Parameters**: `sectionId` - The ID of the section
- **Note**: Now includes complete pagination - finds all pages, not just first 20-50

#### `get_page`
**Purpose**: Get detailed information about a specific page
- **Parameters**: `id` - The page ID

#### `create_page`
**Purpose**: Create a new page in a section with HTML content
- **Parameters**:
  - `sectionId` - The ID of the section
  - `title` - The title of the new page
  - `content` - Optional HTML content for the page

### 🤖 **AI-Powered Tools**

#### `summarize_page`
**Purpose**: Generate AI-powered summary of a specific page's content
- **Parameters**:
  - `pageId` - The ID of the page to summarize
  - `maxLength` - Maximum length of summary (default: 300 words)

#### `summarize_section`
**Purpose**: Generate AI-powered summary of all pages in a section
- **Parameters**:
  - `sectionId` - The ID of the section to summarize
  - `maxPagesPerSection` - Maximum pages to include (default: 10)

## 🧠 **Search Tool Comparison & Best Practices**

### When to Use Which Search Tool:

| Tool | Best For | Searches | Example Use Cases |
|------|----------|----------|-------------------|
| `search_pages` | **Finding specific content anywhere** | All pages across all sections | "Java Townhall", "Leading Agile Teams", "specific meeting notes" |
| `search_by_timeline` | **Time-based content retrieval** | All pages filtered by date ranges | "Q3 2023 notes", "2021 work", "last 6 months" |
| `find_related_content` | **Topic-focused exploration** | Only pages in sections matching the topic | "Java projects", "Azure work", "Microsoft experience" |

### 🔍 **Search Coverage Differences:**

**`search_pages`** - **Most Comprehensive**
- ✅ Searches **ALL 566+ pages** across **ALL sections**
- ✅ No section filtering - finds content anywhere
- ✅ Perfect for finding specific notes when you don't know which section they're in

**`find_related_content`** - **Section-Filtered**  
- ❌ Only searches sections whose **names match your topic**
- ❌ Will miss content in unrelated section names
- ✅ Faster for topic-focused exploration
- ✅ Good for exploring all content related to a major project/client

**Example**: Searching for "Java Townhall"
- `search_pages`: ✅ **Found it** - searches the Company section regardless of name
- `find_related_content`: ❌ **Missed it** - "Java Townhall" doesn't match section name "Company"

### 📅 **Timeline Search Features:**

**Supported Date Formats:**
```
Quarters:     "Q1 2020", "q2 2021", "2023 Q4", "Q3 2023"
Years:        "2020", "2021", "2022"
Months:       "January 2020", "Jan 2020", "2020-01"
Relative:     "recent", "last year", "last 6 months"
```

**Performance Optimizations:**
- ✅ **Full pagination support** - finds ALL pages, not just first 20-50
- ✅ **Sequential API calls** - respects Microsoft Graph rate limits
- ✅ **Precise date filtering** - only returns pages from specified time periods

### 🚀 **Performance & Rate Limiting:**

All search tools now include:
- **Complete pagination** - no missing pages due to API limits
- **Rate limit respect** - sequential API calls prevent throttling
- **Comprehensive coverage** - searches 100% of your content

## Installation

### Installing from npm (Recommended)
```bash
npm install -g mcp-server-onenote
```
The package is now available on the npm registry as of April 27, 2025.

### Running the Package
After installation, you can run the package using:
```bash
mcp-server-onenote
```

Or with npx:
```bash
npx mcp-server-onenote
```

### Installing via Smithery

To install OneNote Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/@modelcontextprotocol/server-onenote):

```bash
npx -y @smithery/cli install @modelcontextprotocol/server-onenote --client claude
```

### Installing from GitHub
```bash
npm install -g github:ZubeidHendricks/azure-onenote-mcp-server
```

## Authentication

This server uses **OAuth2 delegated authentication** with Microsoft Graph API. You'll need to set up Azure environment variables and authenticate once using the built-in auth tool.

### Environment Variables
Set the following environment variables:
- `AZURE_TENANT_ID`: Your Azure tenant ID
- `AZURE_CLIENT_ID`: Your Azure application (client) ID

### Authentication Process
1. **Set environment variables** using the provided scripts
2. **Run the auth tool** to authenticate with Microsoft Graph
3. **The auth tool will**:
   - Open your browser for Microsoft sign-in
   - Handle the OAuth2 flow automatically
   - Save a secure token to `.auth-token` file
   - Token typically expires after 1-2 hours

### Quick Authentication
```bash
# Windows PowerShell
.\set-env.ps1
npm run auth

# Linux/macOS/Git Bash  
source ./set-env.sh
npm run auth
```

### 🔑 **Authentication Flow:**

1. **Initial Setup**: Run auth tool once to authenticate
   ```bash
   npm run auth
   ```
   - Opens browser for Microsoft sign-in
   - User grants consent for OneNote access
   - Saves secure token to `.auth-token` file

2. **Token Management**: 
   - **Automatic**: Server automatically uses saved token
   - **Expiration**: Tokens expire after 1-2 hours for security
   - **Renewal**: Run `npm run auth` again when token expires
   - **Storage**: Token stored locally in `.auth-token` (gitignored)

3. **No Manual Token Handling**: 
   - ❌ No need to copy/paste tokens
   - ❌ No client secrets in configuration  
   - ✅ Secure delegated permissions only
   - ✅ User explicitly controls access

## Using with MCP Client

Add this to your MCP client configuration (e.g. Claude Desktop):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "mcp-server-onenote",
      "env": {
        "AZURE_TENANT_ID": "<YOUR_TENANT_ID>",
        "AZURE_CLIENT_ID": "<YOUR_CLIENT_ID>"
      }
    }
  }
}
```

**Note**: No client secret needed in configuration - authentication is handled via the OAuth2 delegated flow with the auth tool.

## Azure App Registration

### Setup for Delegated Authentication
1. **Go to Azure Portal** and navigate to App registrations
2. **Create a new registration** with these settings:
   - Name: "OneNote MCP Server"
   - Account types: "Accounts in this organizational directory only" (single tenant)
   - Redirect URI: "Web" → `http://localhost:3001/auth/callback`
3. **Add Microsoft Graph API permissions** (Delegated permissions):
   - `Notes.ReadWrite` - Read and write user OneNote notebooks
   - `User.Read` - Sign in and read user profile
   - `offline_access` - Maintain access to data you've given access to
4. **Copy the tenant ID and client ID** for environment configuration
5. **No client secret needed** - delegated auth uses OAuth2 flow with user consent

### 🔐 **Security Benefits of Delegated Auth:**
- ✅ **No stored secrets** - uses secure OAuth2 tokens
- ✅ **User consent required** - user explicitly authorizes access
- ✅ **Token expiration** - automatic security with 1-2 hour token lifetimes
- ✅ **Audit trail** - all actions logged under user's identity

## Development

```bash
# Install dependencies
npm install

# Build
npm run build

# Set environment variables (Windows PowerShell)
npm run env:set

# Set environment variables (Linux/macOS/Git Bash)
npm run env:set-bash

# Run with environment setup (Windows)
npm run dev

# Run locally (after setting environment variables)
npm start

# Run tests
npm test

# Lint
npm run lint
```

### Quick Start Scripts

For convenience, the following scripts are provided to set up your Azure environment variables:

**Windows PowerShell (set env vars and run):**
```powershell
.\set-env-and-run.ps1
```

**Windows Batch (set env vars and run):**
```batch
set-env-and-run.bat
```

**Manual setup - Windows PowerShell:**
```powershell
.\set-env.ps1
npm start
```

**Manual setup - Linux/macOS/Git Bash:**
```bash
source ./set-env.sh
npm start
```

**Using npm scripts:**
```bash
# Windows (automatically sets env vars, builds, and runs)
npm run dev

# Manual setup
npm run env:set      # Windows PowerShell
npm run env:set-bash # Linux/macOS/Git Bash
npm start
```

### 🔐 Secret Management

For security and convenience, the scripts can automatically load your Azure credentials from a `secret.local` file:

1. **Copy the example file:**
   ```bash
   cp secret.local.example secret.local
   ```

2. **Edit `secret.local`** with your actual Azure credentials:
   ```
   AZURE_TENANT_ID=your-actual-tenant-id
   AZURE_CLIENT_ID=your-actual-client-id
   ```

3. **Run the scripts** - they will automatically load from `secret.local`:
   ```bash
   npm run dev  # or .\set-env-and-run.ps1
   ```

**🔒 Security Features:**
- ✅ `secret.local` is in `.gitignore` (never committed to git)
- ✅ Scripts fallback to placeholder values if file doesn't exist
- ✅ Supports comments and empty lines in the secret file
- ✅ Works with all environment setup scripts
- ✅ **No client secrets stored** - uses secure OAuth2 delegated authentication

**⚠️ Security Note**: The `secret.local` file only contains tenant and client IDs (no secrets). Authentication is handled through secure OAuth2 flow.

[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/zubeidhendricks-azure-onenote-mcp-server-badge.png)](https://mseep.ai/app/zubeidhendricks-azure-onenote-mcp-server)


## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for information about contributing to this repository.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

## Changelog

### 0.1.0 (2025-04-27)
- Initial release
- Core functionality for OneNote notebook, section, and page management
- Published to npm registry

### 0.1.1 (2025-04-27)
- Added executable bin to package.json
- Fixed issue where npx command couldn't determine executable to run
- Updated README with clear running instructions
