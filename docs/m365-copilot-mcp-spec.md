# Microsoft 365 Copilot MCP Server Specification

**Version:** 1.0.0  
**Status:** Draft  
**Date:** January 2026

---

## Executive Summary

### What This Server Does

An MCP server that brings **Microsoft 365 enterprise data** into AI assistant workflows (GitHub Copilot, Claude Desktop). Query SharePoint, OneDrive, email, calendar, Teams meetings—all through natural language, with full permission enforcement.

### Tools at a Glance

| Tool | API | What It Does | Best For |
|------|-----|--------------|----------|
| `m365_retrieve` | Retrieval | Returns text chunks for YOUR AI to reason over | Custom RAG, deep analysis |
| `m365_chat` | Chat | M365 Copilot synthesizes an answer | Quick Q&A, people/calendar |
| `m365_meetings` | Meeting Insights | AI summaries, action items, mentions | Post-meeting follow-up |
| `m365_search` | Search | Semantic document discovery | Finding files |

### API Comparison

| Capability | Chat | Retrieval | Meetings | Search |
|------------|------|-----------|----------|--------|
| Natural language Q&A | ✅ | ❌ chunks | ❌ | ❌ |
| Custom AI reasoning | ❌ | ✅ **Your LLM** | ❌ | ✅ **Your LLM** |
| SharePoint/OneDrive | ✅ | ✅ | ❌ | OneDrive only |
| Email/Calendar/Teams | ✅ | ❌ | Meetings | ❌ |
| KQL filtering | ❌ | ✅ | ❌ | Limited |
| Multi-turn | ✅ | ❌ | ❌ | ❌ |

---

## 1. Overview

An MCP (Model Context Protocol) server that enables AI assistants (GitHub Copilot, Claude Desktop) to interact with Microsoft 365 Copilot via the Chat API. This brings enterprise data grounding—SharePoint, OneDrive, email, calendar, Teams—directly into AI assistant workflows while maintaining Microsoft 365's security, compliance, and permissions model.

### 1.1 Value Proposition

| Benefit | Description |
|---------|-------------|
| **Enterprise Grounding** | Access to Microsoft 365 semantic index (SharePoint, OneDrive, Exchange, Teams) |
| **Permission-Aware** | Respects user-level and tenant-level Microsoft 365 permissions |
| **Web + Work** | Combines enterprise search grounding with optional web search |
| **Compliance Built-In** | Data never leaves Microsoft 365 trust boundary |
| **No RAG Infrastructure** | Eliminates need for custom vector indexes or data pipelines |

### 1.2 Key Differentiators from Deep Research

| Aspect | Deep Research | M365 Copilot |
|--------|---------------|---------------|
| **Data Source** | Public web | Enterprise Microsoft 365 data + optional web |
| **Auth Model** | Azure AI Foundry (Cognitive Services) | Microsoft Graph API (delegated permissions) |
| **Use Cases** | General research, market analysis | Enterprise Q&A, document search, meeting context |
| **Grounding** | Web search only | Enterprise search + file context + web |

### 1.3 Example Use Cases by Persona

| Persona | Primary Tools | Example Queries |
|---------|---------------|------------------|
| **Chief Architect** | `m365_retrieve`, `m365_meetings` | "Find ADRs about microservices", "Action items from security review" |
| **Sales** | `m365_chat`, `m365_meetings` | "Emails with Contoso about renewal", "Prep for tomorrow's call" |
| **Legal/Compliance** | `m365_retrieve`, `m365_search` | "Find contracts mentioning liability caps", "Policy docs from 2024" |
| **Project Manager** | `m365_meetings`, `m365_chat` | "Outstanding action items from sprint meetings", "Who owns the API?" |
| **Executive** | `m365_chat`, `m365_meetings` | "Summarize this week's board prep", "What decisions need my input?" |

## 2. Architecture

### 2.1 High-Level Flow

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  MCP Client     │     │   MCP Server    │     │  Microsoft      │
│  (GH Copilot/   │────▶│  m365-copilot   │────▶│  Graph API      │
│   Claude)       │     │                 │     │  /beta/copilot  │
└─────────────────┘     └─────────────────┘     └─────────────────┘
                                │                        │
                                │                        ▼
                                │               ┌─────────────────┐
                                │               │  M365 Copilot   │
                                │               │  Chat Service   │
                                │               │  (Enterprise +  │
                                │               │   Web Search)   │
                                │               └─────────────────┘
                                │
                                ▼
                        ┌─────────────────┐
                        │   Conversation  │
                        │     Store       │
                        │  (In-Memory)    │
                        └─────────────────┘
```

### 2.2 API Endpoints Used

| API | Endpoint | Purpose |
|-----|----------|---------|
| **Chat API** | `POST /beta/copilot/conversations` | Create conversation |
| | `POST /beta/copilot/conversations/{id}/chat` | Sync chat |
| | `POST /beta/copilot/conversations/{id}/chatOverStream` | Streaming chat (SSE) |
| **Retrieval API** | `POST /beta/copilot/retrieval` | Get text chunks for RAG |
| **Search API** | `POST /beta/copilot/search` | Semantic document search |
| **Meeting Insights** | `GET /v1.0/copilot/users/{id}/onlineMeetings/{id}/aiInsights` | List meeting insights |
| | `GET /v1.0/copilot/users/{id}/onlineMeetings/{id}/aiInsights/{id}` | Get specific insight |

### 2.3 Component Architecture

```python
# Core components
server.py              # MCP server with tools + health/root endpoints
clients/
  ├── chat.py          # Chat API client (conversations)
  ├── retrieval.py     # Retrieval API client (RAG chunks)
  ├── search.py        # Search API client (document discovery)
  └── meetings.py      # Meeting Insights API client
auth.py                # Authentication (delegated user auth via Graph)
conversation.py        # Conversation state management
```

### 2.4 HTTP Endpoints (non-MCP)

```python
@mcp.custom_route("/", methods=["GET"])
async def root_info(request):
    """Service discovery endpoint."""
    return JSONResponse({
        "service": "m365-copilot-mcp",
        "status": "running",
        "mcp_endpoint": "/mcp",
        "health_endpoint": "/health",
    })

@mcp.custom_route("/health", methods=["GET"])
async def health_check(request):
    """Health check for Container Apps / K8s probes."""
    return JSONResponse({"status": "healthy"})
```

### 2.5 When to Use Each API

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         API SELECTION DECISION TREE                          │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  "I want to ask a question and get an answer"                               │
│       └── Use Chat API (m365_chat)                                          │
│           • M365 Copilot synthesizes answer                                 │
│           • Includes email, calendar, Teams, SharePoint, OneDrive           │
│           • Multi-turn conversation support                                 │
│                                                                              │
│  "I want raw context for MY AI to reason over"                              │
│       └── Use Retrieval API (m365_retrieve)                                 │
│           • Returns text chunks with relevance scores                       │
│           • You control the reasoning/synthesis                             │
│           • SharePoint, OneDrive, Copilot Connectors                        │
│           • Supports KQL filtering                                          │
│                                                                              │
│  "I want to find documents (not their content)"                             │
│       └── Use Search API (m365_search)                                      │
│           • Semantic + lexical hybrid search                                │
│           • Returns file metadata, URLs, previews                           │
│           • OneDrive only (currently)                                       │
│                                                                              │
│  "I want meeting summaries and action items"                                │
│       └── Use Meeting Insights API (m365_meetings)                          │
│           • Pre-generated AI summaries                                      │
│           • Structured action items with owners                             │
│           • Mention tracking (when you were mentioned)                      │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

## 3. Authentication

### 3.1 Required Microsoft Graph Permissions (Delegated)

All permissions are delegated (user context only—no application permissions supported).

#### Chat API Permissions
| Permission | Purpose |
|------------|---------|
| `Sites.Read.All` | SharePoint site access |
| `Mail.Read` | Email content access |
| `People.Read.All` | People and contacts |
| `OnlineMeetingTranscript.Read.All` | Meeting transcripts |
| `Chat.Read` | Teams chat history |
| `ChannelMessage.Read.All` | Teams channel messages |
| `ExternalItem.Read.All` | Copilot connectors content |

> **Note:** Chat API requires ALL permissions simultaneously.

#### Retrieval API Permissions
| Permission | Purpose |
|------------|---------|
| `Files.Read.All` | OneDrive file access |
| `Sites.Read.All` | SharePoint content |
| `ExternalItem.Read.All` | Copilot connectors (optional) |

#### Search API Permissions
| Permission | Purpose |
|------------|---------|
| `Files.Read.All` | OneDrive file search |
| `Sites.Read.All` | SharePoint search |

#### Meeting Insights API Permissions
| Permission | Purpose |
|------------|---------|
| `OnlineMeetingTranscript.Read.All` | Access meeting AI insights |
| `OnlineMeeting.Read` | List user's meetings |

### 3.2 Authentication Methods (Priority Order)

| Method | Use Case | Configuration |
|--------|----------|---------------|
| **Interactive Browser** | Local dev, first-time setup | Default for new users |
| **Device Code Flow** | Headless/SSH environments | Fallback when browser unavailable |
| **Cached Token** | Subsequent runs | Tokens cached in `~/.m365-copilot-mcp/` |
| **Client Credentials + User Assertion** | Advanced scenarios | On-behalf-of flow |

### 3.3 App Registration Requirements

```
Azure AD App Registration:
├── Redirect URI: http://localhost:8400 (for auth callback)
├── Platform: Mobile and desktop applications
├── API Permissions: (see 3.1 above)
├── Supported account types: Single tenant or Multitenant
└── Token configuration: Access tokens + Refresh tokens
```

### 3.4 Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_CLIENT_ID` | Yes | App registration client ID |
| `AZURE_TENANT_ID` | Yes | Azure AD tenant ID |
| `AZURE_CLIENT_SECRET` | No | Only for confidential clients |
| `M365_COPILOT_TIMEOUT` | No | Request timeout (default: 60s) |
| `M365_COPILOT_CACHE_DIR` | No | Token cache location |

## 4. MCP Tools

### Tool Selection Guide

Choose the right tool based on what you need:

| If you need... | Use | Why |
|----------------|-----|-----|
| Quick answer from enterprise data | `m365_chat` | M365 synthesizes, includes email/calendar |
| Raw text chunks for custom analysis | `m365_retrieve` | You control reasoning, KQL filtering |
| Meeting summaries & action items | `m365_meetings` | Structured insights from transcripts |
| Find documents by topic | `m365_search` | Discovery when you don't know file names |
| Analyze specific known files | `m365_chat_with_files` | Ground answers in particular documents |

**Escalation pattern:** `m365_chat` (quick) → `m365_retrieve` (deep) → `m365_chat_with_files` (specific)

---

### 4.1 Tool: `m365_retrieve` (Retrieval API)

**Best for:** Deep research where you want YOUR AI to reason over enterprise data.

```python
@mcp.tool()
async def m365_retrieve(
    query: str = Field(
        description="Natural language query for enterprise content. Be specific: include document types, projects, or topics. E.g., 'Q4 revenue projections for ACME deal' not just 'revenue'."
    ),
    data_source: Literal["sharepoint", "onedrive", "connectors"] = Field(
        default="sharepoint",
        description="Where to search: 'sharepoint' (team sites, wikis), 'onedrive' (personal files), 'connectors' (external systems via Copilot connectors)"
    ),
    filter_expression: str | None = Field(
        default=None,
        description="Optional KQL filter to narrow scope. Examples: 'path:https://contoso.sharepoint.com/sites/HR', 'FileType:pdf', 'LastModifiedTime>2024-01-01'"
    ),
    max_results: int = Field(
        default=25,
        description="Number of text chunks to return (1-25). More chunks = more context but longer processing."
    ),
    ctx: Context = None,
) -> CallToolResult:
    """Retrieve raw text chunks from M365 for YOUR AI to reason over. Returns relevance-scored excerpts from SharePoint/OneDrive—you control synthesis. Use for: custom analysis, cross-document reasoning, when you need source text not just answers. Use m365_chat instead for: quick Q&A, calendar/email questions, when M365's answer is sufficient."""
```

**API Details:**
- Endpoint: `POST /beta/copilot/retrieval`
- Returns: Text chunks with `relevanceScore` (0-1), source URLs, metadata
- Rate limit: 200 requests/user/hour
- Permissions: `Files.Read.All`, `Sites.Read.All`, `ExternalItem.Read.All`

---

### 4.2 Tool: `m365_chat` (Chat API)

**Best for:** Quick Q&A where M365 Copilot synthesizes the answer.

```python
@mcp.tool()
async def m365_chat(
    message: str = Field(
        description="Question for M365 Copilot. Works best for: people lookup, calendar queries, email summaries, quick factual questions about enterprise data. E.g., 'Who owns budget approval?' or 'Summarize emails from Contoso this week'."
    ),
    conversation_id: str | None = Field(
        default=None,
        description="For follow-up questions, pass the conversation_id from previous response. Omit to start fresh."
    ),
    web_search: bool = Field(
        default=True,
        description="Include public web in grounding. Set False for sensitive/internal-only queries."
    ),
    ctx: Context = None,
) -> CallToolResult:
    """Quick Q&A with M365 Copilot. Gets synthesized answers from email, calendar, Teams, SharePoint, OneDrive. Use for: people questions, meeting schedules, email summaries, enterprise facts. Supports multi-turn conversation. Use m365_retrieve instead when: you need raw source text, want to control reasoning, or need cross-document analysis."""
```

---

### 4.3 Tool: `m365_meetings` (Meeting Insights API)

**Best for:** Extracting structured insights from Teams meetings.

```python
@mcp.tool()
async def m365_meetings(
    meeting_id: str | None = Field(
        default=None,
        description="Teams meeting ID (from calendar or meeting URL). Omit to list recent meetings."
    ),
    join_url: str | None = Field(
        default=None,
        description="Full Teams join URL as alternative to meeting_id. E.g., 'https://teams.microsoft.com/l/meetup-join/...'"
    ),
    since: str | None = Field(
        default=None,
        description="ISO datetime to filter meetings from. E.g., '2026-01-06T00:00:00Z' for last week. Defaults to 7 days ago if omitted."
    ),
    ctx: Context = None,
) -> CallToolResult:
    """Get AI-generated meeting summaries, action items, and mentions from Teams. Returns structured data: notes, decisions, tasks with owners, when you were mentioned. Requires: transcription enabled, ~4hr after meeting ends. Use for: post-meeting follow-up, finding action items, checking what you missed. Does NOT work for: channel meetings, meetings without transcription."""
```

**API Details:**
- List meetings: `GET /copilot/users/{userId}/onlineMeetings/{meetingId}/aiInsights`
- Get insight: `GET /copilot/users/{userId}/onlineMeetings/{meetingId}/aiInsights/{insightId}`
- Returns: `meetingNotes[]`, `actionItems[]`, `viewpoint.mentionEvents[]`
- Permissions: `OnlineMeetingTranscript.Read.All`

**Response Structure:**
```json
{
  "meetingNotes": [
    {
      "title": "Project Status Discussion",
      "text": "Team reviewed timeline and resource allocation...",
      "subpoints": [
        {"title": "Timeline", "text": "On track for Q2 delivery..."},
        {"title": "Decision", "text": "Proceed with Option B pending budget approval"}
      ]
    }
  ],
  "actionItems": [
    {
      "title": "Budget Review",
      "text": "Complete budget review by Friday",
      "ownerDisplayName": "Sarah Chen"
    }
  ],
  "viewpoint": {
    "mentionEvents": [
      {
        "eventDateTime": "2026-01-10T14:30:00Z",
        "transcriptUtterance": "We need to get approval from [You] before proceeding",
        "speaker": {"displayName": "John Smith"}
      }
    ]
  }
}
```

---

### 4.4 Tool: `m365_search` (Search API)

**Best for:** Discovering relevant documents when you don't know exact location.

```python
@mcp.tool()
async def m365_search(
    query: str = Field(
        description="What documents to find. Use natural language—semantic search handles synonyms. E.g., 'Q3 board presentation' or 'contracts with renewal clauses'."
    ),
    path_filter: str | None = Field(
        default=None,
        description="Scope to OneDrive folder path. E.g., '/Documents/Projects/Alpha' to search only that folder."
    ),
    page_size: int = Field(
        default=25,
        description="Results to return (1-100). Start with 25, increase if needed."
    ),
    ctx: Context = None,
) -> CallToolResult:
    """Find documents in OneDrive using semantic + keyword hybrid search. Returns file metadata, previews, URLs—not full content. Use for: discovering relevant files, finding documents by topic when you don't know exact names. Use m365_retrieve instead when: you need actual document content, want text for analysis. Limitation: OneDrive only (SharePoint search coming)."""
```

**API Details:**
- Endpoint: `POST /beta/copilot/search`
- Returns: File metadata, preview text, URLs
- Rate limit: 200 requests/user/hour
- Permissions: `Files.Read.All`, `Sites.Read.All`

---

### 4.5 Tool: `m365_chat_with_files` (Chat API + File Context)

**Best for:** Asking questions about specific documents you already know.

```python
@mcp.tool()
async def m365_chat_with_files(
    message: str = Field(
        description="Question about the provided files. E.g., 'Summarize key risks' or 'Compare revenue projections between these reports'."
    ),
    file_uris: list[str] = Field(
        description="SharePoint/OneDrive file URIs to analyze. Get URIs from m365_search results or SharePoint URLs. E.g., ['https://contoso.sharepoint.com/sites/Sales/proposal.docx']"
    ),
    conversation_id: str | None = Field(
        default=None,
        description="For follow-up questions about same files, pass conversation_id from previous response."
    ),
    ctx: Context = None,
) -> CallToolResult:
    """Ask questions about specific documents you already have URIs for. M365 Copilot reads the files and answers. Use for: summarizing known documents, comparing specific files, extracting info from a particular doc. Use m365_search first when: you need to find the files. Use m365_retrieve when: you want raw text chunks, not Copilot's synthesis."""
```

## 5. Conversation Management

### 5.1 Multi-Turn Conversations

The server maintains conversation state to support multi-turn interactions:

```python
@dataclass
class ConversationState:
    id: str
    created_at: datetime
    turn_count: int
    last_activity: datetime
    display_name: str  # Set to first message
```

**Conversation lifecycle:**
1. First message creates new conversation (returns `conversation_id`)
2. Subsequent messages include `conversation_id` for context continuity
3. Conversations expire after 1 hour of inactivity (server-side cleanup)
4. Client can start fresh by omitting `conversation_id`

### 5.2 Conversation Store

```python
class ConversationStore:
    """In-memory conversation state with TTL-based cleanup."""
    
    def create(self) -> str: ...
    def get(self, id: str) -> ConversationState | None: ...
    def update_activity(self, id: str) -> None: ...
    def cleanup_expired(self) -> None: ...
```

## 6. Response Format

### 6.1 Standard Response Structure

```python
@dataclass
class M365CopilotResponse:
    text: str                          # Main answer text
    conversation_id: str               # For multi-turn
    turn_count: int                    # Conversation turn number
    attributions: list[Attribution]    # Source citations
    sensitivity_label: SensitivityLabel | None
```

### 6.2 Attribution Types

| Type | Description |
|------|-------------|
| `annotation` | Inline entity reference (person, event, file) |
| `citation` | Source document reference |
| `grounding` | Enterprise search match |

### 6.3 Citation Format in Responses

Citations appear as:
- Inline: `[^1^]`, `[^2^]` style markers
- Entity tags: `<Person>John Doe</Person>`, `<Event>Standup</Event>`, `<File>Report.docx</File>`
- URLs: Deep links to source content

## 7. Error Handling

### 7.1 Error Categories

| Category | HTTP Status | Handling |
|----------|-------------|----------|
| Auth failure | 401 | Trigger re-authentication flow |
| Permission denied | 403 | User lacks required M365 permissions |
| Rate limited | 429 | Exponential backoff retry |
| Gateway timeout | 504 | Retry with shorter query (known limitation) |
| Service error | 5xx | Retry with backoff |

### 7.2 Known Limitations

1. **No Actions**: Chat API is read-only—cannot create files, send emails, or schedule meetings
2. **Text Only**: Responses are text—no image generation or code interpreter
3. **Timeout Risk**: Long-running queries may hit gateway timeouts
4. **Single-Turn Web Toggle**: Disabling web search must be done per-message
5. **Beta API**: Subject to breaking changes

## 8. Configuration

### 8.1 VS Code / GitHub Copilot (`mcp.json`)

```json
{
  "servers": {
    "m365-copilot": {
      "type": "stdio",
      "command": "uvx",
      "args": ["--from", "git+https://github.com/org/m365-copilot-mcp.git", "m365-copilot"],
      "env": {
        "PYTHONUNBUFFERED": "1",
        "AZURE_CLIENT_ID": "your-app-client-id",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

### 8.2 Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "m365-copilot": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/org/m365-copilot-mcp.git", "m365-copilot"],
      "env": {
        "AZURE_CLIENT_ID": "your-app-client-id",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

### 8.3 Development Mode (HTTP)

```bash
# Run as HTTP server for debugging
m365-copilot --http --port 8000
```

## 9. Licensing Requirements

| Requirement | Details |
|-------------|---------|
| **User License** | Microsoft 365 Copilot add-on license required |
| **Tenant** | Microsoft 365 commercial tenant |
| **Region** | Global (GCC/GCC-High/DoD not currently supported) |

> **Cost**: Chat API usage is included with Microsoft 365 Copilot license—no additional API costs.

## 10. Security Considerations

### 10.1 Data Handling

- **No Data Storage**: Server does not persist user data or responses
- **Token Security**: Refresh tokens cached locally with user-only permissions
- **Query Logging**: Queries truncated in logs (GDPR compliance)

### 10.2 Permission Model

```
User Permission Scope:
├── Can only query data they have access to in Microsoft 365
├── SharePoint permissions enforced
├── Email permissions enforced
├── Teams permissions enforced
└── Sensitivity labels preserved in responses
```

## 11. Implementation Decisions (ADRs)

### ADR-001: Use Microsoft Graph Beta API

**Context**: Chat API only available in `/beta` endpoint.  
**Decision**: Use beta API with version pinning.  
**Consequences**: Must monitor for breaking changes; cannot use in production-critical scenarios without mitigation.

### ADR-002: Delegated Auth Only

**Context**: Chat API requires user context—no application-only support.  
**Decision**: Implement interactive browser auth with token caching.  
**Consequences**: Requires user sign-in; cannot run as unattended service.

### ADR-003: In-Memory Conversation State

**Context**: Need multi-turn conversation support.  
**Decision**: Store conversation IDs in-memory with TTL cleanup.  
**Consequences**: Conversations lost on server restart; acceptable for MCP stdio pattern.

### ADR-004: Streaming via SSE

**Context**: Chat API supports both sync and streaming endpoints.  
**Decision**: Use streaming (`chatOverStream`) by default for better UX.  
**Consequences**: More complex response parsing; better perceived latency.

### ADR-005: No File Upload

**Context**: API accepts file URIs, not file content.  
**Decision**: Require users to provide SharePoint/OneDrive URIs.  
**Consequences**: Files must already be in Microsoft 365; cannot analyze local files.

### ADR-006: Hybrid SDK Strategy

**Context**: `msgraph-sdk` doesn't support SSE streaming; Chat API uses SSE.  
**Decision**: Use `msgraph-sdk` for Retrieval/Search/Meetings APIs; use `httpx` + `httpx-sse` for Chat API streaming.  
**Consequences**: Two HTTP client patterns; auth token shared via `azure-identity`.

### ADR-007: Markdown Output Format

**Context**: Need consistent output format across tools.  
**Decision**: Return markdown with inline citations (`[^1^]`) matching deep-research pattern.  
**Consequences**: Requires parsing API JSON and formatting; better UX consistency.

## 12. Project Structure

**Repository:** `github.com/renepajta/m365-copilot-mcp`

```
m365-copilot-mcp/
├── src/
│   └── m365_copilot/
│       ├── __init__.py
│       ├── server.py              # MCP server, tool definitions
│       ├── auth.py                # Authentication flows
│       ├── conversation.py        # Conversation state management
│       └── clients/
│           ├── __init__.py
│           ├── base.py            # Base Graph client
│           ├── chat.py            # Chat API (m365_chat, m365_chat_with_files)
│           ├── retrieval.py       # Retrieval API (m365_retrieve)
│           ├── search.py          # Search API (m365_search)
│           └── meetings.py        # Meeting Insights API (m365_meetings)
├── tests/
│   ├── test_server.py
│   ├── test_auth.py
│   └── clients/
│       ├── test_chat.py
│       ├── test_retrieval.py
│       ├── test_search.py
│       └── test_meetings.py
├── docs/
│   └── m365-copilot-mcp-spec.md
├── pyproject.toml
├── README.md
└── LICENSE
```

## 13. Dependencies

```toml
[project]
dependencies = [
    "mcp>=1.25.0",
    "azure-identity>=1.26.0",
    "msgraph-sdk>=1.0.0",           # Graph SDK for Retrieval/Search/Meetings
    "httpx>=0.27.0",                 # Async HTTP for Chat API streaming
    "httpx-sse>=0.4.0",              # SSE support for streaming chat
    "pydantic>=2.12.5",
    "python-dotenv>=1.0.0",
]

[project.scripts]
m365-copilot = "m365_copilot.server:main"
```

## 14. Alignment with Deep Research MCP

This section tracks alignment with the deep-research MCP server patterns.

### 14.1 Implementation Patterns to Adopt

| Pattern | Deep Research | M365 Copilot | Notes |
|---------|---------------|--------------|-------|
| Health endpoint | `@mcp.custom_route("/health")` | **TODO** | Required for Container Apps probes |
| Root info endpoint | `@mcp.custom_route("/")` | **TODO** | Service discovery |
| Request ID logging | `gen_request_id()` → 6 hex chars | **TODO** | Log correlation |
| Query truncation | `truncate_query(query, 50)` | **TODO** | GDPR compliance |
| Usage stats | `UsageStats` dataclass | **TODO** | Token counting, latency |
| Error returns | `CallToolResult(isError=True)` | **TODO** | Proper MCP error signaling |
| Progress reporting | `ctx.report_progress(pct, 100, msg)` | **TODO** | Long-running operations |
| Retry with backoff | `MAX_RETRIES=3, RETRY_DELAYS=[5,15,30]` | **TODO** | Transient failures |
| Noisy logger silencing | `logging.getLogger("httpx").setLevel(WARNING)` | **TODO** | Clean logs |

### 14.2 Server Entry Point Pattern

```python
def main():
    """Run the MCP server."""
    import argparse
    parser = argparse.ArgumentParser(description="M365 Copilot MCP Server")
    parser.add_argument("--http", action="store_true", help="Run as HTTP server")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()

    if args.http:
        mcp.settings.port = args.port
        logger.info("Starting HTTP server on port %d", args.port)
        mcp.run(transport="streamable-http", mount_path="/mcp")
    else:
        logger.info("Starting stdio server for VS Code MCP")
        mcp.run(transport="stdio")
```

### 14.3 pyproject.toml Entry Point

```toml
[project.scripts]
m365-copilot = "m365_copilot.server:main"
```

## 15. Design Decisions (Resolved)

| # | Decision | Resolution |
|---|----------|------------|
| **D-1** | Streaming behavior | **Accumulate** — SSE with `httpx-sse`, collect all chunks, return complete text |
| **D-2** | Token refresh | **Auto-refresh** — Use `azure-identity` credential with built-in refresh |
| **D-3** | Timeout values | **Per-tool**: chat=60s, retrieve=90s, search=60s, meetings=30s, chat_with_files=120s |
| **D-4** | Meeting date filter | **Add `since` param** — List recent meetings, then fetch insights |
| **D-5** | Progress reporting | **Yes** — Report at 25%, 50%, 75% for long operations |
| **D-6** | Conversation TTL | **1 hour** — Align with M365 API, cleanup on server |
| **D-7** | Output format | **Markdown with citations** — `According to [^1^]...` style, matches deep-research |
| **D-8** | Sensitivity labels | **Footer** — `---\n⚠️ Sensitivity: {label}` at end of response |
| **D-9** | Graph client library | **Hybrid** — `msgraph-sdk` for Retrieval/Search/Meetings, `httpx-sse` for Chat streaming |
| **D-10** | Multi-file limit | **Document in tool** — API supports multiple URIs, test for practical limit |

## 17. Future Considerations

| Feature | Priority | Notes |
|---------|----------|-------|
| Agent ID support | Medium | Target specific Copilot agents |
| Copilot Retrieval API expansion | High | SharePoint in Search API when available |
| Interaction Export API | Medium | Audit trail of your Copilot usage |
| Change Notifications API | Low | Real-time Copilot interaction monitoring |
| Batch operations | Low | Multiple queries in single call |

## 16. API Rate Limits Summary

| API | Rate Limit | Notes |
|-----|------------|-------|
| Chat API | Not documented | Subject to Graph throttling |
| Retrieval API | 200 req/user/hour | Max 25 results per request |
| Search API | 200 req/user/hour | Max 100 results per request |
| Meeting Insights | Standard Graph limits | Available ~4 hours post-meeting |

## 18. Development Setup

### 18.1 Prerequisites

- WSL2 (Ubuntu recommended)
- Python 3.11+
- Azure CLI (`az login` for local dev)
- Microsoft 365 Copilot license (for testing)

### 18.2 Repository Setup

```bash
# User creates folder, then:
cd /path/to/m365-copilot-mcp

# Initialize repo
git init
git remote add origin git@github.com:renepajta/m365-copilot-mcp.git

# Create virtual environment
python3 -m venv .venv
source .venv/bin/activate

# Install in editable mode (after pyproject.toml created)
pip install -e ".[dev]"
```

### 18.3 Environment Variables

```bash
# .env file (not committed)
AZURE_CLIENT_ID=your-app-client-id
AZURE_TENANT_ID=your-tenant-id
# AZURE_CLIENT_SECRET only if using confidential client
```

### 18.4 Running Locally

```bash
# stdio mode (for VS Code MCP testing)
m365-copilot

# HTTP mode (for debugging)
m365-copilot --http --port 8000
```

### 18.5 Testing

```bash
pytest tests/ -v
```

## 19. References

- [Microsoft 365 Copilot APIs Overview](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/copilot-apis-overview)
- [Chat API Overview](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/chat/overview)
- [Retrieval API Overview](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/retrieval/overview)
- [Search API Overview](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/search/overview)
- [Meeting AI Insights API](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/meeting-transcripts/meeting-insights)
- [Create Copilot Conversations](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/chat/copilotroot-post-conversations)
- [Synchronous Chat Endpoint](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/chat/copilotconversation-chat)
- [Streamed Chat Endpoint](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/chat/copilotconversation-chatoverstream)
- [Microsoft Graph Permissions Reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [MCP Specification](https://modelcontextprotocol.io)
