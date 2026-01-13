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
   - `OnlineMeeting.Read`
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

### 3. Configure

Create a `.env` file:

```bash
AZURE_CLIENT_ID=your-app-client-id
AZURE_TENANT_ID=your-tenant-id
```

### 4. Test Locally

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

On first run, you'll be prompted to authenticate:

1. **Browser auth** (default): Opens browser for Microsoft sign-in
2. **Device code** (fallback): For headless environments, displays a code to enter at microsoft.com/devicelogin

Tokens are cached locally in `~/.m365-copilot-mcp/` for subsequent runs.

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_CLIENT_ID` | Yes | App registration client ID |
| `AZURE_TENANT_ID` | Yes | Azure AD tenant ID |
| `AZURE_CLIENT_SECRET` | No | Only for confidential clients |
| `M365_COPILOT_TIMEOUT` | No | Request timeout in seconds (default: 60) |
| `M365_COPILOT_CACHE_DIR` | No | Token cache location |

## Troubleshooting

### "Insufficient permissions"
- Ensure admin consent is granted for all API permissions
- Verify user has Microsoft 365 Copilot license

### "Token expired"
- Delete `~/.m365-copilot-mcp/` and re-authenticate

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
