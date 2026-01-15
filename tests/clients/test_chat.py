"""Tests for Chat API client."""

import pytest
from unittest.mock import AsyncMock, MagicMock, patch
import json

from m365_copilot.clients.chat import (
    ChatClient,
    ChatResponse,
    ChatApiError,
)
from m365_copilot.clients.base import Attribution


class TestChatResponse:
    """Tests for ChatResponse dataclass."""

    def test_to_markdown_simple(self):
        """Should format simple response."""
        response = ChatResponse(
            text="This is the answer.",
            conversation_id="conv-123",
        )
        
        result = response.to_markdown()
        assert "This is the answer." in result

    def test_to_markdown_with_citations(self):
        """Should include citations section."""
        response = ChatResponse(
            text="According to [^1^], the answer is 42.",
            conversation_id="conv-123",
            attributions=[
                Attribution(
                    type="citation",
                    text="Source doc",
                    url="https://example.com/doc",
                    title="The Source",
                )
            ],
        )
        
        result = response.to_markdown()
        assert "Sources:" in result
        assert "[^1^]" in result
        assert "https://example.com/doc" in result

    def test_to_markdown_with_sensitivity(self):
        """Should include sensitivity label."""
        response = ChatResponse(
            text="Confidential info here.",
            conversation_id="conv-123",
            sensitivity_label="Confidential",
        )
        
        result = response.to_markdown()
        assert "⚠️ Sensitivity: Confidential" in result


class TestChatClient:
    """Tests for ChatClient."""

    @pytest.fixture
    def mock_credential(self):
        """Create mock credential."""
        cred = MagicMock()
        cred.get_token.return_value = MagicMock(token="test-token")
        return cred

    @pytest.mark.asyncio
    async def test_create_conversation_success(self, mock_credential):
        """Should create conversation and return ID."""
        # Patch the SDK client before ChatClient instantiation
        with patch(
            "m365_copilot.auth.create_sdk_client"
        ) as mock_sdk_class:
            # Create mock SDK client instance
            mock_sdk = MagicMock()
            mock_sdk_class.return_value = mock_sdk
            
            # Mock the conversation creation
            mock_result = MagicMock()
            mock_result.id = "new-conv-123"
            mock_sdk.copilot.conversations.post = AsyncMock(return_value=mock_result)
            
            client = ChatClient(mock_credential)
            result = await client.create_conversation()
            
            assert result == "new-conv-123"
            mock_sdk.copilot.conversations.post.assert_called_once()

    @pytest.mark.asyncio
    async def test_create_conversation_failure(self, mock_credential):
        """Should raise ChatApiError on failure."""
        with patch(
            "m365_copilot.auth.create_sdk_client"
        ) as mock_sdk_class:
            mock_sdk = MagicMock()
            mock_sdk_class.return_value = mock_sdk
            mock_sdk.copilot.conversations.post = AsyncMock(
                side_effect=Exception("API error")
            )
            
            client = ChatClient(mock_credential)
            
            with pytest.raises(ChatApiError):
                await client.create_conversation()
