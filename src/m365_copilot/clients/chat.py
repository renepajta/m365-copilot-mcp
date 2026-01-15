"""Chat API client for M365 Copilot.

Implements the Chat API using the official Microsoft SDK (microsoft-agents-m365copilot-beta).
Falls back to streaming endpoint if available.

Endpoints:
- POST /copilot/conversations - Create conversation
- POST /copilot/conversations/{id}/microsoft.graph.copilot.chat - Synchronous chat
- POST /copilot/conversations/{id}/microsoft.graph.copilot.chatOverStream - Streaming (SSE)
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

import httpx
from httpx_sse import aconnect_sse, SSEError

from microsoft_agents_m365copilot_beta import (
    AgentsM365CopilotBetaServiceClient,
)
from microsoft_agents_m365copilot_beta.generated.models.copilot_conversation import (
    CopilotConversation,
)
from microsoft_agents_m365copilot_beta.generated.models.copilot_conversation_location import (
    CopilotConversationLocation,
)
from microsoft_agents_m365copilot_beta.generated.models.copilot_conversation_request_message_parameter import (
    CopilotConversationRequestMessageParameter,
)
from microsoft_agents_m365copilot_beta.generated.copilot.conversations.item.microsoft_graph_copilot_chat.chat_post_request_body import (
    ChatPostRequestBody,
)

from m365_copilot.clients.base import (
    Attribution,
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


class ChatClient:
    """Client for M365 Copilot Chat API using official Microsoft SDK."""

    BETA_BASE_URL = "https://graph.microsoft.com/beta"

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        self.credential = credential
        self.timeout = timeout or CHAT_TIMEOUT
        
        # Create SDK client with correct beta API configuration
        from m365_copilot.auth import create_sdk_client
        self._sdk_client = create_sdk_client(credential)

    async def _get_access_token(self) -> str:
        """Get access token from credential."""
        from m365_copilot.auth import GRAPH_SCOPES
        token = self.credential.get_token(*GRAPH_SCOPES)
        return token.token

    async def create_conversation(self) -> str:
        """Create a new conversation and return its ID.

        Returns:
            Conversation ID from M365 Copilot.
        """
        request_id = gen_request_id()

        try:
            # Use SDK to create conversation
            new_conversation = CopilotConversation()
            result = await self._sdk_client.copilot.conversations.post(new_conversation)
            
            if result is None or result.id is None:
                raise ChatApiError("Failed to create conversation: no ID returned")
            
            conversation_id = result.id
            logger.info("[%s] Created conversation: %s", request_id, conversation_id)
            return conversation_id
            
        except Exception as e:
            logger.error("[%s] Failed to create conversation: %s", request_id, str(e))
            raise ChatApiError(f"Failed to create conversation: {e}")

    async def chat(
        self,
        conversation_id: str,
        message: str,
        *,
        web_search: bool = True,
        file_uris: list[str] | None = None,
        request_id: str | None = None,
    ) -> ChatResponse:
        """Send a chat message and get response.

        Uses official SDK with proper model serialization.
        Falls back to streaming endpoint for better UX if SDK fails.

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

        logger.info(
            "[%s] Chat: %s (web=%s, files=%d)",
            request_id,
            truncate_query(message),
            web_search,
            len(file_uris) if file_uris else 0,
        )

        # Try SDK-based synchronous endpoint first (proper model serialization)
        try:
            return await self._chat_sdk(
                conversation_id, message, web_search, file_uris, request_id
            )
        except Exception as e:
            logger.warning(
                "[%s] SDK chat failed, trying streaming fallback: %s",
                request_id,
                str(e),
            )
            # Fall back to streaming endpoint
            token = await self._get_access_token()
            return await self._chat_streaming(
                conversation_id, message, token, web_search, file_uris, request_id
            )

    async def _chat_sdk(
        self,
        conversation_id: str,
        message: str,
        web_search: bool,
        file_uris: list[str] | None,
        request_id: str,
    ) -> ChatResponse:
        """Use official SDK for chat - handles timezone format correctly."""
        
        # Build request body using SDK models
        request_body = ChatPostRequestBody()
        
        # Set message using the proper SDK model
        msg_param = CopilotConversationRequestMessageParameter()
        msg_param.text = message
        request_body.message = msg_param
        
        # Set location hint with timezone
        # The SDK model serializes this correctly as "timeZone" (camelCase)
        location = CopilotConversationLocation()
        location.time_zone = "America/Los_Angeles"
        request_body.location_hint = location
        
        # Add grounding options via additional_data if web search disabled
        if not web_search:
            request_body.additional_data["groundingOptions"] = {"disableWebGrounding": True}
        
        # Add file context via additional_data if provided
        if file_uris:
            request_body.additional_data["externalContexts"] = [
                {"type": "fileUri", "value": uri} for uri in file_uris
            ]

        # Call the SDK chat endpoint
        result = await self._sdk_client.copilot.conversations.by_copilot_conversation_id(
            conversation_id
        ).microsoft_graph_copilot_chat.post(request_body)
        
        if result is None:
            raise ChatApiError("Chat returned no response")
        
        # Parse response from SDK result
        text = ""
        attributions: list[Attribution] = []
        sensitivity_label: str | None = None
        turn_count = result.turn_count or 1
        
        # Extract assistant response from messages
        # Messages list contains: [user_message, assistant_response]
        # We want the last message which is the assistant's response
        if result.messages and len(result.messages) > 0:
            # Get the last message (assistant response)
            # Usually messages[0] is echo of user input, messages[1] is assistant response
            assistant_msg = result.messages[-1]
            
            # Get text from the message
            if hasattr(assistant_msg, 'text') and assistant_msg.text:
                text = assistant_msg.text
            
            # Get attributions
            if hasattr(assistant_msg, 'attributions') and assistant_msg.attributions:
                for attr in assistant_msg.attributions:
                    attributions.append(
                        Attribution(
                            type=getattr(attr, 'type', 'citation') or 'citation',
                            text=getattr(attr, 'text', '') or '',
                            url=getattr(attr, 'url', None),
                            title=getattr(attr, 'title', None),
                        )
                    )
            
            # Get sensitivity label
            if hasattr(assistant_msg, 'sensitivity_label') and assistant_msg.sensitivity_label:
                sl = assistant_msg.sensitivity_label
                if hasattr(sl, 'display_name') and sl.display_name:
                    sensitivity_label = sl.display_name
        
        logger.info(
            "[%s] Chat (SDK) complete: %d chars, %d citations",
            request_id,
            len(text),
            len(attributions),
        )

        return ChatResponse(
            text=text,
            conversation_id=conversation_id,
            turn_count=turn_count,
            attributions=attributions,
            sensitivity_label=sensitivity_label,
        )

    async def _chat_streaming(
        self,
        conversation_id: str,
        message: str,
        token: str,
        web_search: bool,
        file_uris: list[str] | None,
        request_id: str,
    ) -> ChatResponse:
        """Stream response using SSE endpoint (fallback)."""
        url = f"{self.BETA_BASE_URL}/copilot/conversations/{conversation_id}/microsoft.graph.copilot.chatOverStream"

        # Build body using SDK-compatible format
        body: dict[str, Any] = {
            "message": {
                "text": message,
            },
            "locationHint": {
                "timeZone": "America/Los_Angeles",
            }
        }
        
        if not web_search:
            body["groundingOptions"] = {"disableWebGrounding": True}
        
        if file_uris:
            body["externalContexts"] = [
                {"type": "fileUri", "value": uri} for uri in file_uris
            ]

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
            "[%s] Chat (streaming) complete: %d chars, %d citations",
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
