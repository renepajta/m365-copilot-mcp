"""MCP Server for Microsoft 365 Copilot APIs.

Exposes 5 tools for enterprise data access:
- m365_retrieve: Get text chunks for RAG (Retrieval API)
- m365_chat: Quick Q&A with M365 Copilot (Chat API)
- m365_meetings: Meeting summaries and action items (Meeting Insights API)
- m365_search: Document discovery (Search API)
- m365_chat_with_files: Ask questions about specific files (Chat API + files)
"""

from __future__ import annotations

import argparse
import logging
import os
from datetime import datetime, timezone
from typing import Literal

from dotenv import load_dotenv
from mcp.server.fastmcp import Context, FastMCP
from mcp.types import CallToolResult, TextContent
from pydantic import Field
from starlette.responses import JSONResponse

from m365_copilot.auth import get_credential
from m365_copilot.clients.base import gen_request_id, truncate_query
from m365_copilot.clients.chat import ChatClient
from m365_copilot.clients.meetings import MeetingsClient
from m365_copilot.clients.retrieval import RetrievalClient
from m365_copilot.clients.search import SearchClient
from m365_copilot.conversation import get_conversation_store

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# Silence noisy loggers
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

# Initialize MCP server
mcp = FastMCP(
    name="m365-copilot",
    instructions="""Microsoft 365 Copilot MCP Server.

Access enterprise data from SharePoint, OneDrive, Email, Calendar, and Teams meetings.
All queries respect user-level Microsoft 365 permissions.

Available tools:
- m365_retrieve: Get raw text chunks for YOUR AI to analyze (RAG)
- m365_chat: Quick Q&A where M365 Copilot synthesizes answers
- m365_meetings: Get meeting summaries, action items, and mentions
- m365_search: Find documents by topic/content
- m365_chat_with_files: Ask questions about specific SharePoint/OneDrive files

Authentication: Requires Azure AD app with delegated permissions and M365 Copilot license.
""",
)

# Global clients (initialized on first use)
_credential = None
_chat_client: ChatClient | None = None
_retrieval_client: RetrievalClient | None = None
_search_client: SearchClient | None = None
_meetings_client: MeetingsClient | None = None


def _get_credential():
    """Get or create authentication credential."""
    global _credential
    if _credential is None:
        _credential = get_credential()
    return _credential


def _get_chat_client() -> ChatClient:
    """Get or create Chat API client."""
    global _chat_client
    if _chat_client is None:
        _chat_client = ChatClient(_get_credential())
    return _chat_client


def _get_retrieval_client() -> RetrievalClient:
    """Get or create Retrieval API client."""
    global _retrieval_client
    if _retrieval_client is None:
        _retrieval_client = RetrievalClient(_get_credential())
    return _retrieval_client


def _get_search_client() -> SearchClient:
    """Get or create Search API client."""
    global _search_client
    if _search_client is None:
        _search_client = SearchClient(_get_credential())
    return _search_client


def _get_meetings_client() -> MeetingsClient:
    """Get or create Meetings API client."""
    global _meetings_client
    if _meetings_client is None:
        _meetings_client = MeetingsClient(_get_credential())
    return _meetings_client


# =============================================================================
# HTTP Endpoints (non-MCP)
# =============================================================================


@mcp.custom_route("/", methods=["GET"])
async def root_info(request):
    """Service discovery endpoint."""
    return JSONResponse({
        "service": "m365-copilot-mcp",
        "version": "0.1.0",
        "status": "running",
        "mcp_endpoint": "/mcp",
        "health_endpoint": "/health",
        "description": "MCP server for Microsoft 365 Copilot APIs",
    })


@mcp.custom_route("/health", methods=["GET"])
async def health_check(request):
    """Health check for Container Apps / K8s probes."""
    return JSONResponse({
        "status": "healthy",
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })


# =============================================================================
# MCP Tools
# =============================================================================


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
    """Retrieve raw text chunks from M365 for YOUR AI to reason over.

    Returns relevance-scored excerpts from SharePoint/OneDrive—you control synthesis.

    Use for:
    - Custom analysis and cross-document reasoning
    - When you need source text, not just answers
    - Deep research where you want control over synthesis

    Use m365_chat instead for:
    - Quick Q&A where M365's answer is sufficient
    - Calendar/email questions
    - People lookup
    """
    request_id = gen_request_id()
    logger.info("[%s] m365_retrieve: %s", request_id, truncate_query(query))

    try:
        # Report progress
        if ctx:
            await ctx.report_progress(25, 100, "Connecting to M365...")

        client = _get_retrieval_client()

        if ctx:
            await ctx.report_progress(50, 100, "Retrieving content...")

        response = await client.retrieve(
            query,
            data_source=data_source,
            filter_expression=filter_expression,
            max_results=max_results,
            request_id=request_id,
        )

        if ctx:
            await ctx.report_progress(100, 100, "Complete")

        return CallToolResult(
            content=[TextContent(type="text", text=response.to_markdown())],
            isError=False,
        )

    except Exception as e:
        logger.error("[%s] m365_retrieve error: %s", request_id, e)
        return CallToolResult(
            content=[TextContent(type="text", text=f"Error retrieving content: {e}")],
            isError=True,
        )


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
    """Quick Q&A with M365 Copilot.

    Gets synthesized answers from email, calendar, Teams, SharePoint, OneDrive.
    Supports multi-turn conversation.

    Use for:
    - People questions ('Who owns X?')
    - Meeting schedules and availability
    - Email summaries
    - Enterprise facts and policies

    Use m365_retrieve instead when:
    - You need raw source text
    - You want to control reasoning
    - You need cross-document analysis
    """
    request_id = gen_request_id()
    logger.info("[%s] m365_chat: %s", request_id, truncate_query(message))

    try:
        if ctx:
            await ctx.report_progress(25, 100, "Connecting to M365 Copilot...")

        client = _get_chat_client()
        store = get_conversation_store()

        # Create or get conversation
        if conversation_id:
            conv_state = store.get(conversation_id)
            if not conv_state:
                # Conversation expired, create new one
                api_conversation_id = await client.create_conversation()
                conv_state = store.create(display_name=message[:50])
            else:
                api_conversation_id = conversation_id
        else:
            api_conversation_id = await client.create_conversation()
            conv_state = store.create(display_name=message[:50])

        if ctx:
            await ctx.report_progress(50, 100, "Processing query...")

        response = await client.chat(
            api_conversation_id,
            message,
            web_search=web_search,
            request_id=request_id,
        )

        # Update conversation state
        turn = conv_state.increment_turn()
        response.turn_count = turn
        response.conversation_id = conv_state.id

        if ctx:
            await ctx.report_progress(100, 100, "Complete")

        # Add conversation metadata to response
        output = response.to_markdown()
        output += f"\n\n---\n*Conversation: `{conv_state.id}` (turn {turn})*"

        return CallToolResult(
            content=[TextContent(type="text", text=output)],
            isError=False,
        )

    except Exception as e:
        logger.error("[%s] m365_chat error: %s", request_id, e)
        return CallToolResult(
            content=[TextContent(type="text", text=f"Error in chat: {e}")],
            isError=True,
        )


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
    """Get AI-generated meeting summaries, action items, and mentions from Teams.

    Returns structured data: notes, decisions, tasks with owners, when you were mentioned.

    Requires:
    - Transcription enabled during meeting
    - ~4 hours after meeting ends for insights to be ready

    Use for:
    - Post-meeting follow-up
    - Finding action items assigned to you
    - Checking what you missed in meetings

    Does NOT work for:
    - Channel meetings
    - Meetings without transcription enabled
    """
    request_id = gen_request_id()
    logger.info("[%s] m365_meetings: id=%s", request_id, meeting_id or "list")

    try:
        if ctx:
            await ctx.report_progress(25, 100, "Connecting to Teams...")

        client = _get_meetings_client()

        if ctx:
            await ctx.report_progress(50, 100, "Fetching meeting data...")

        # If no meeting_id provided, list recent meetings
        if not meeting_id and not join_url:
            since_dt = None
            if since:
                since_dt = datetime.fromisoformat(since.replace("Z", "+00:00"))

            meetings = await client.list_meetings(since=since_dt, request_id=request_id)

            if ctx:
                await ctx.report_progress(100, 100, "Complete")

            if not meetings:
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text="No meetings found in the specified time range."
                    )],
                    isError=False,
                )

            output = "# Recent Meetings\n\n"
            output += "Select a meeting ID to get AI insights:\n\n"
            for meeting in meetings:
                output += meeting.to_markdown() + "\n"

            return CallToolResult(
                content=[TextContent(type="text", text=output)],
                isError=False,
            )

        # Get insights for specific meeting
        insight = await client.get_insights(
            meeting_id or "",
            join_url=join_url,
            request_id=request_id,
        )

        if ctx:
            await ctx.report_progress(100, 100, "Complete")

        return CallToolResult(
            content=[TextContent(type="text", text=insight.to_markdown())],
            isError=False,
        )

    except Exception as e:
        logger.error("[%s] m365_meetings error: %s", request_id, e)
        return CallToolResult(
            content=[TextContent(type="text", text=f"Error getting meeting insights: {e}")],
            isError=True,
        )


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
    """Find documents in OneDrive using semantic + keyword hybrid search.

    Returns file metadata, previews, URLs—not full content.

    Use for:
    - Discovering relevant files
    - Finding documents by topic when you don't know exact names
    - Building a list of files to analyze

    Use m365_retrieve instead when:
    - You need actual document content
    - You want text for analysis

    Limitation: OneDrive only (SharePoint search coming)
    """
    request_id = gen_request_id()
    logger.info("[%s] m365_search: %s", request_id, truncate_query(query))

    try:
        if ctx:
            await ctx.report_progress(25, 100, "Searching OneDrive...")

        client = _get_search_client()

        if ctx:
            await ctx.report_progress(50, 100, "Processing results...")

        response = await client.search(
            query,
            path_filter=path_filter,
            page_size=page_size,
            request_id=request_id,
        )

        if ctx:
            await ctx.report_progress(100, 100, "Complete")

        return CallToolResult(
            content=[TextContent(type="text", text=response.to_markdown())],
            isError=False,
        )

    except Exception as e:
        logger.error("[%s] m365_search error: %s", request_id, e)
        return CallToolResult(
            content=[TextContent(type="text", text=f"Error searching: {e}")],
            isError=True,
        )


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
    """Ask questions about specific documents you already have URIs for.

    M365 Copilot reads the files and answers.

    Use for:
    - Summarizing known documents
    - Comparing specific files
    - Extracting info from particular docs

    Use m365_search first when:
    - You need to find the files

    Use m365_retrieve when:
    - You want raw text chunks, not Copilot's synthesis
    """
    request_id = gen_request_id()
    logger.info(
        "[%s] m365_chat_with_files: %s (%d files)",
        request_id,
        truncate_query(message),
        len(file_uris),
    )

    try:
        if ctx:
            await ctx.report_progress(25, 100, "Connecting to M365 Copilot...")

        client = _get_chat_client()
        store = get_conversation_store()

        # Create or get conversation
        if conversation_id:
            conv_state = store.get(conversation_id)
            if not conv_state:
                api_conversation_id = await client.create_conversation()
                conv_state = store.create(display_name=message[:50])
            else:
                api_conversation_id = conversation_id
        else:
            api_conversation_id = await client.create_conversation()
            conv_state = store.create(display_name=message[:50])

        if ctx:
            await ctx.report_progress(50, 100, "Analyzing files...")

        response = await client.chat_with_files(
            api_conversation_id,
            message,
            file_uris,
            request_id=request_id,
        )

        turn = conv_state.increment_turn()
        response.turn_count = turn
        response.conversation_id = conv_state.id

        if ctx:
            await ctx.report_progress(100, 100, "Complete")

        output = response.to_markdown()
        output += f"\n\n---\n*Conversation: `{conv_state.id}` (turn {turn})*"
        output += f"\n*Files analyzed: {len(file_uris)}*"

        return CallToolResult(
            content=[TextContent(type="text", text=output)],
            isError=False,
        )

    except Exception as e:
        logger.error("[%s] m365_chat_with_files error: %s", request_id, e)
        return CallToolResult(
            content=[TextContent(type="text", text=f"Error analyzing files: {e}")],
            isError=True,
        )


# =============================================================================
# Entry Point
# =============================================================================


def main():
    """Run the MCP server."""
    parser = argparse.ArgumentParser(description="M365 Copilot MCP Server")
    parser.add_argument(
        "--http",
        action="store_true",
        help="Run as HTTP server (for debugging)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="HTTP server port (default: 8000)",
    )
    parser.add_argument(
        "--auth",
        action="store_true",
        help="Authenticate interactively and save credentials for future use",
    )
    args = parser.parse_args()

    # Validate required environment variables
    if not os.getenv("AZURE_CLIENT_ID"):
        logger.warning("AZURE_CLIENT_ID not set - authentication will fail")
    if not os.getenv("AZURE_TENANT_ID"):
        logger.warning("AZURE_TENANT_ID not set - authentication will fail")

    # Handle one-time authentication
    if args.auth:
        from m365_copilot.auth import authenticate_and_save
        print("🔐 Starting interactive authentication...")
        print("   A browser window will open. Sign in with your M365 account.")
        print()
        try:
            record = authenticate_and_save()
            print(f"✅ Authenticated as: {record.username}")
            print(f"   Credentials saved. Future runs will use cached tokens.")
        except Exception as e:
            print(f"❌ Authentication failed: {e}")
            raise SystemExit(1)
        return

    if args.http:
        logger.info("Starting HTTP server on port %d", args.port)
        mcp.settings.port = args.port
        mcp.run(transport="streamable-http", mount_path="/mcp")
    else:
        logger.info("Starting stdio server for VS Code MCP")
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
