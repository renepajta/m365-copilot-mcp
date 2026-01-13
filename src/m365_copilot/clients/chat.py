"""Chat API client for M365 Copilot.

Implements the Chat API using httpx-sse for streaming responses (ADR-004, ADR-006).

Endpoints:
- POST /beta/copilot/conversations - Create conversation
- POST /beta/copilot/conversations/{id}/chatOverStream - Streaming chat (SSE)
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, AsyncIterator

import httpx
from httpx_sse import aconnect_sse

from m365_copilot.clients.base import (
    Attribution,
    GraphClient,
    format_citations,
    format_sensitivity_label,
    gen_request_id,
    truncate_query,
)

if TYPE_CHECKING:
    from azure.core.credentials import TokenCredential

logger = logging.getLogger(__name__)

# Chat API timeout (longer for complex queries)
CHAT_TIMEOUT = 60
CHAT_WITH_FILES_TIMEOUT = 120


@dataclass
class ChatResponse:
    """Response from M365 Copilot Chat API."""

    text: str
    conversation_id: str
    turn_count: int = 1
    attributions: list[Attribution] = field(default_factory=list)
    sensitivity_label: str | None = None

    def to_markdown(self) -> str:
        """Format response as markdown with citations."""
        parts = [self.text]

        citations = format_citations(self.attributions)
        if citations:
            parts.append(citations)

        sensitivity = format_sensitivity_label(self.sensitivity_label)
        if sensitivity:
            parts.append(sensitivity)

        return "\n".join(parts)


class ChatClient(GraphClient):
    """Client for M365 Copilot Chat API with SSE streaming."""

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        super().__init__(credential, timeout=timeout or CHAT_TIMEOUT)

    async def create_conversation(self) -> str:
        """Create a new conversation and return its ID.

        Returns:
            Conversation ID from M365 Copilot.
        """
        request_id = gen_request_id()
        url = f"{self.BETA_BASE_URL}/copilot/conversations"

        response = await self._make_request(
            "POST",
            url,
            json={},
            request_id=request_id,
        )

        if response.status_code != 201:
            logger.error(
                "[%s] Failed to create conversation: %d %s",
                request_id,
                response.status_code,
                response.text,
            )
            raise ChatApiError(
                f"Failed to create conversation: {response.status_code}"
            )

        data = response.json()
        conversation_id = data.get("id")
        logger.info("[%s] Created conversation: %s", request_id, conversation_id)
        return conversation_id

    async def chat(
        self,
        conversation_id: str,
        message: str,
        *,
        web_search: bool = True,
        file_uris: list[str] | None = None,
        request_id: str | None = None,
    ) -> ChatResponse:
        """Send a chat message and get streaming response.

        Uses SSE endpoint for streaming, accumulates all chunks (ADR-004).

        Args:
            conversation_id: ID of existing conversation.
            message: User message to send.
            web_search: Include web search in grounding.
            file_uris: SharePoint/OneDrive file URIs to include as context.
            request_id: Request ID for logging.

        Returns:
            ChatResponse with full accumulated text and attributions.
        """
        request_id = request_id or gen_request_id()
        url = f"{self.BETA_BASE_URL}/copilot/conversations/{conversation_id}/chatOverStream"

        logger.info(
            "[%s] Chat: %s (web=%s, files=%d)",
            request_id,
            truncate_query(message),
            web_search,
            len(file_uris) if file_uris else 0,
        )

        # Build request body
        body: dict[str, Any] = {
            "messages": [
                {
                    "content": message,
                    "role": "user",
                }
            ],
        }

        # Add grounding options
        if not web_search:
            body["groundingOptions"] = {"disableWebGrounding": True}

        # Add file context if provided
        if file_uris:
            body["externalContexts"] = [
                {"type": "fileUri", "value": uri} for uri in file_uris
            ]

        # Get access token
        token = await self._get_access_token()

        # Stream response using SSE
        text_chunks: list[str] = []
        attributions: list[Attribution] = []
        sensitivity_label: str | None = None

        async with httpx.AsyncClient(
            timeout=httpx.Timeout(
                CHAT_WITH_FILES_TIMEOUT if file_uris else CHAT_TIMEOUT
            )
        ) as client:
            async with aconnect_sse(
                client,
                "POST",
                url,
                json=body,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                    "Accept": "text/event-stream",
                    "X-Request-Id": request_id,
                },
            ) as event_source:
                async for event in event_source.aiter_sse():
                    if event.event == "copilotMessageDelta":
                        # Parse delta content
                        try:
                            data = json.loads(event.data)
                            delta = data.get("delta", {})

                            # Accumulate text
                            if "content" in delta:
                                text_chunks.append(delta["content"])

                            # Collect attributions
                            if "attributions" in delta:
                                for attr in delta["attributions"]:
                                    attributions.append(
                                        Attribution(
                                            type=attr.get("type", "citation"),
                                            text=attr.get("text", ""),
                                            url=attr.get("url"),
                                            title=attr.get("title"),
                                        )
                                    )

                            # Check for sensitivity label
                            if "sensitivityLabel" in delta:
                                sensitivity_label = delta["sensitivityLabel"].get(
                                    "displayName"
                                )

                        except json.JSONDecodeError:
                            logger.warning(
                                "[%s] Failed to parse SSE event: %s",
                                request_id,
                                event.data[:100],
                            )

                    elif event.event == "copilotMessageComplete":
                        # Final event - may contain final attributions
                        try:
                            data = json.loads(event.data)
                            if "attributions" in data:
                                for attr in data["attributions"]:
                                    # Dedupe by URL
                                    if not any(
                                        a.url == attr.get("url") for a in attributions
                                    ):
                                        attributions.append(
                                            Attribution(
                                                type=attr.get("type", "citation"),
                                                text=attr.get("text", ""),
                                                url=attr.get("url"),
                                                title=attr.get("title"),
                                            )
                                        )
                        except json.JSONDecodeError:
                            pass

                    elif event.event == "error":
                        logger.error("[%s] SSE error: %s", request_id, event.data)
                        raise ChatApiError(f"Chat error: {event.data}")

        full_text = "".join(text_chunks)
        logger.info(
            "[%s] Chat complete: %d chars, %d citations",
            request_id,
            len(full_text),
            len(attributions),
        )

        return ChatResponse(
            text=full_text,
            conversation_id=conversation_id,
            attributions=attributions,
            sensitivity_label=sensitivity_label,
        )

    async def chat_with_files(
        self,
        conversation_id: str,
        message: str,
        file_uris: list[str],
        *,
        request_id: str | None = None,
    ) -> ChatResponse:
        """Chat with specific files as context.

        Convenience wrapper for chat() with file_uris.
        """
        return await self.chat(
            conversation_id,
            message,
            file_uris=file_uris,
            request_id=request_id,
        )


class ChatApiError(Exception):
    """Error from M365 Copilot Chat API."""

    pass
