# OneNote MCP Server - Delegated Authentication Branch

This branch implements **delegated authentication** with user sign-in instead of client credentials flow.

## 🔄 Key Differences from Main Branch

| Feature | Main Branch | Delegated Branch |
|---------|-------------|------------------|
| **Authentication** | Client Credentials | Delegated (User Sign-in) |
| **Admin Consent** | Required | Not Required |
| **User Interaction** | None | Device Code Flow |
| **Permissions** | Application | Delegated |
| **API Endpoints** | `/users/{id}/...` | `/me/...` |
| **Client Secret** | Required | Not Required |

## 🚀 Quick Start

### 1. Azure App Registration Setup

1. Go to Azure Portal → App Registrations → Your App
2. **Authentication** → Add platform → Web:
   - **Redirect URI**: `http://localhost:3001/auth/callback`
   - **Front-channel logout URL**: (leave empty)
3. **API Permissions** → Add **Delegated permissions**:
   - `Notes.ReadWrite` - Read and write user OneNote notebooks
   - `User.Read` - Sign in and read user profile
4. **No admin consent required!** ✅

### 2. Configuration

Create `secret.local` file:
```
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
# No AZURE_CLIENT_SECRET needed!
```

### 3. Run the Server

```bash
npm run dev
```

The server will prompt for user authentication:
```
🔐 USER AUTHENTICATION REQUIRED
================================
Opening browser for Microsoft sign-in...
🌐 Local auth server started at http://localhost:3001
🌐 Opening browser to: https://login.microsoftonline.com/...
⏳ Waiting for you to complete sign-in in your browser...
```

## 🔐 Authentication Flow

1. **First Run**: Local web server + browser authentication
2. **Subsequent Runs**: Token cached automatically 
3. **Token Refresh**: Handled automatically when expired
4. **User Context**: All API calls use `/me` endpoints (signed-in user)
5. **Enterprise-Friendly**: No device code flow - works with corporate restrictions

## ✅ Advantages

- **✅ No Admin Consent Required** - Works immediately
- **✅ User Controls Access** - User grants permissions directly
- **✅ Simpler Azure Setup** - Just delegated permissions
- **✅ Token Caching** - Smooth subsequent usage
- **✅ Cross-Platform** - Works on all operating systems

## ⚠️ Considerations

- **❌ Requires User Interaction** - Not fully automated
- **❌ Per-User Authentication** - Each user must sign in
- **❌ Internet Connection Required** - For initial sign-in
- **❌ Token Expiration** - May require periodic re-authentication

## 🛠️ Tools Available

All the same OneNote tools as the main branch:
- `list_notebooks` - List all user's notebooks
- `get_notebook` - Get notebook details  
- `create_notebook` - Create new notebook
- `list_sections` - List sections in notebook
- `create_section` - Create new section
- `list_pages` - List pages in section
- `get_page` - Get page details
- `create_page` - Create new page with content

## 🎯 Perfect For

- **Development & Testing** - No admin approval needed
- **Personal Use** - Access your own OneNote data
- **Prototyping** - Quick setup without IT involvement
- **Demo & Learning** - Easy to try out

## 🔄 Switching Between Branches

```bash
# Switch to main branch (client credentials)
git checkout main

# Switch to delegated branch (user sign-in)  
git checkout delegated
```

## 📚 More Information

- **Main Branch**: Enterprise-ready with client credentials
- **This Branch**: User-friendly with interactive sign-in
- Both branches provide the same OneNote functionality!
