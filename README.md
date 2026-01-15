# m365-copilot-mcp

MCP server for Microsoft 365 Copilot APIs — bring enterprise data into GitHub Copilot & Claude Desktop.

## What This Does

Query SharePoint, OneDrive, email, calendar, and Teams meetings through natural language—all with full Microsoft 365 permission enforcement.

| Tool | API | What It Does | Best For |
|------|-----|--------------|----------|
| `m365_retrieve` | Retrieval | Returns text chunks for YOUR AI to reason over | Custom RAG, deep analysis |
| `m365_chat` | Chat | M365 Copilot synthesizes an answer | Quick Q&A, people/calendar |
| `m365_meetings` | Meeting Insights | AI summaries, action items, mentions | Post-meeting follow-up |
| `m365_search` | Search | Semantic document discovery | Finding files |
| `m365_chat_with_files` | Chat + Files | Ask questions about specific documents | Summarizing known files |

## Requirements

- **Microsoft 365 Copilot license** (required for API access)
- Python 3.11+
- Azure AD app registration with delegated permissions

## Quick Start

### 1. Create Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Name: `m365-copilot-mcp`
4. Supported account types: Single tenant (or Multitenant)
5. Redirect URI: `http://localhost:8400` (Mobile and desktop applications)
6. Add API permissions (Delegated):
   - `Sites.Read.All`
   - `Mail.Read`
   - `People.Read.All`
   - `OnlineMeetingTranscript.Read.All`
   - `Chat.Read`
   - `ChannelMessage.Read.All`
   - `ExternalItem.Read.All`
   - `Files.Read.All`
   - `OnlineMeetings.Read`
7. Grant admin consent

### 2. Install

```bash
# Clone the repo
git clone https://github.com/renepajta/m365-copilot-mcp.git
cd m365-copilot-mcp

# Create virtual environment
python3 -m venv .venv
source .venv/bin/activate

# Install
pip install -e ".[dev]"
```

### 2a. WSL-Specific Setup (Windows Subsystem for Linux)

If running in WSL, additional setup is required for browser-based authentication:

#### Enable WSL Interoperability

WSL interoperability allows Linux to launch Windows executables (like browsers). Check if it's enabled:

```bash
ls /proc/sys/fs/binfmt_misc/WSLInterop
```

If the file doesn't exist, check your `/etc/wsl.conf`:

```bash
cat /etc/wsl.conf
```

Ensure interop is enabled (or not explicitly disabled):

```ini
[interop]
enabled = true
appendWindowsPath = true
```

After editing, restart WSL from **Windows PowerShell**:

```powershell
wsl --shutdown
```

Then reopen your WSL terminal.

#### Install wslu (WSL Utilities)

[wslu](https://wslutiliti.es/wslu/) provides `wslview` which opens Windows browsers from WSL:

```bash
# Ubuntu/Debian
sudo apt install wslu

# Other distros: see https://wslutiliti.es/wslu/install.html
```

Verify installation:

```bash
wslview https://google.com
```

This should open Google in your Windows default browser.

#### Alternative: Run from Windows

If WSL interop doesn't work in your environment, you can configure VS Code to use Windows Python instead of WSL Python in your `.vscode/mcp.json`:

```json
{
  "servers": {
    "m365-copilot": {
      "type": "stdio",
      "command": "C:\\path\\to\\m365-copilot-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "m365_copilot.server"],
      "env": {
        "PYTHONUNBUFFERED": "1",
        "AZURE_CLIENT_ID": "your-app-client-id",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

### 3. Configure

Create a `.env` file:

```bash
AZURE_CLIENT_ID=your-app-client-id
AZURE_TENANT_ID=your-tenant-id

# Optional: specify account when multiple are cached
# AZURE_USERNAME=user@contoso.com
```

### 4. Authenticate (One-Time)

Run the authentication command once to cache your credentials:

```bash
m365-copilot --auth
```

A browser window will open. Sign in with your M365 account. Your credentials are saved to `~/.m365-copilot-mcp/` for future use.

### 5. Test Locally

```bash
# Run in HTTP mode for debugging
m365-copilot --http --port 8000

# Check health
curl http://localhost:8000/health
```

## MCP Client Configuration

### VS Code / GitHub Copilot

Add to your `.vscode/mcp.json`:

```json
{
  "servers": {
    "m365-copilot": {
      "type": "stdio",
      "command": "uvx",
      "args": ["--from", "git+https://github.com/renepajta/m365-copilot-mcp.git", "m365-copilot"],
      "env": {
        "PYTHONUNBUFFERED": "1",
        "AZURE_CLIENT_ID": "your-app-client-id",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

### Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "m365-copilot": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/renepajta/m365-copilot-mcp.git", "m365-copilot"],
      "env": {
        "AZURE_CLIENT_ID": "your-app-client-id",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

## Example Queries

```
# Quick Q&A (m365_chat)
"Who owns the budget approval process?"
"Summarize emails from Contoso this week"
"What meetings do I have tomorrow?"

# Deep research (m365_retrieve)
"Find all ADRs related to microservices architecture"
"Get policy documents about data retention from 2024"

# Meeting follow-up (m365_meetings)
"What action items came out of the security review?"
"When was I mentioned in yesterday's standup?"

# Document discovery (m365_search)
"Find contracts mentioning liability caps"
"Q3 board presentation slides"

# File analysis (m365_chat_with_files)
"Summarize key risks in this proposal"
"Compare revenue projections between these two reports"
```

## Authentication Flow

This MCP server uses **delegated permissions**, meaning it accesses Microsoft 365 data **on behalf of you**, the signed-in user.

### How It Works

```
┌─────────────┐     ┌─────────────────┐     ┌──────────────┐     ┌─────────────┐
│   VS Code   │     │   MCP Server    │     │  Microsoft   │     │   Graph     │
│   Copilot   │     │  (subprocess)   │     │  Entra ID    │     │   API       │
└──────┬──────┘     └────────┬────────┘     └──────┬───────┘     └──────┬──────┘
       │                     │                     │                    │
       │ 1. Start server     │                     │                    │
       ├────────────────────►│                     │                    │
       │                     │                     │                    │
       │ 2. Tool call        │                     │                    │
       ├────────────────────►│                     │                    │
       │                     │                     │                    │
       │                     │ 3. Opens browser    │                    │
       │◄─ ─ ─ ─ ─ ─ ─ ─ ─ ─►│    for sign-in     │                    │
       │                     │                     │                    │
       │ 4. YOU sign in ─────────────────────────►│                    │
       │                     │                     │                    │
       │                     │ 5. Token (for YOU)  │                    │
       │                     │◄────────────────────┤                    │
       │                     │                     │                    │
       │                     │ 6. API call with YOUR token             │
       │                     ├─────────────────────────────────────────►│
       │                     │                     │                    │
       │                     │ 7. YOUR data only                       │
       │ 8. Results          │◄─────────────────────────────────────────┤
       │◄────────────────────┤                     │                    │
```

### Key Security Points

| Aspect | Explanation |
|--------|-------------|
| **App Registration (SPN)** | Defines what permissions *can* be requested—not whose data is accessed |
| **User Sign-in** | Required on first use; you authenticate with your M365 account |
| **Access Token** | Contains YOUR identity; grants access only to YOUR mailbox, files, etc. |
| **Token Cache** | Stored locally in `~/.m365-copilot-mcp/` for subsequent runs |

### Authentication Methods

1. **Interactive Browser** (default): Opens your browser for Microsoft sign-in
2. **Device Code** (fallback): For headless/SSH environments—displays a code to enter at [microsoft.com/devicelogin](https://microsoft.com/devicelogin)

### First-Time Setup

**Recommended:** Run `m365-copilot --auth` once before using with VS Code or Claude Desktop. This caches your credentials so the MCP server can authenticate silently.

If you skip this step, you'll see on first tool use:

- **Browser flow**: A browser window opens → Sign in with your M365 account → Consent to permissions
- **Device code flow**: A message in VS Code Output panel with a code → Visit the URL → Enter the code → Sign in

After authentication, tokens are cached and subsequent requests don't require re-authentication (until token expires, typically 1-24 hours depending on tenant policy).

### Why Delegated Permissions?

This approach is more secure for personal developer tools:

| Delegated (This Server) | Application Permissions |
|------------------------|------------------------|
| ✅ Requires user sign-in | ❌ No user sign-in needed |
| ✅ Access only YOUR data | ⚠️ Access ANY user's data |
| ✅ User or admin consent | ⚠️ Admin consent required |
| ✅ Ideal for personal tools | Better for background services |

The SPN cannot access any Microsoft 365 data without you explicitly signing in first.

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_CLIENT_ID` | Yes | App registration client ID |
| `AZURE_TENANT_ID` | Yes | Azure AD tenant ID |
| `AZURE_CLIENT_SECRET` | No | Only for confidential clients |
| `M365_COPILOT_TIMEOUT` | No | Request timeout in seconds (default: 60) |
| `M365_COPILOT_CACHE_DIR` | No | Token cache location |

## Troubleshooting

### WSL: "WSL Interoperability is disabled"
- Check `/etc/wsl.conf` for `[interop] enabled = false` and remove/change it
- Restart WSL: `wsl --shutdown` from Windows PowerShell
- Verify with: `ls /proc/sys/fs/binfmt_misc/WSLInterop`

### WSL: Browser doesn't open
- Install wslu: `sudo apt install wslu`
- Test: `wslview https://google.com`
- Alternative: Run `python login.py` from Windows PowerShell

### "Insufficient permissions"
- Ensure admin consent is granted for all API permissions
- Verify user has Microsoft 365 Copilot license

### "Token expired"
- Run `m365-copilot --auth` to re-authenticate
- Or delete `~/.m365-copilot-mcp/` and re-authenticate

### "Gateway timeout" on long queries
- Break complex queries into smaller parts
- Use `m365_retrieve` for better control over scope

## API Rate Limits

| API | Rate Limit | Notes |
|-----|------------|-------|
| Chat API | Not documented | Subject to Graph throttling |
| Retrieval API | 200 req/user/hour | Max 25 results per request |
| Search API | 200 req/user/hour | Max 100 results per request |
| Meeting Insights | Standard Graph limits | Available ~4 hours post-meeting |

## Development

```bash
# Run tests
pytest tests/ -v

# Lint
ruff check src/

# Type check
mypy src/
```

## License

MIT

## References

- [Microsoft 365 Copilot APIs Overview](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/copilot-apis-overview)
- [MCP Specification](https://modelcontextprotocol.io)
