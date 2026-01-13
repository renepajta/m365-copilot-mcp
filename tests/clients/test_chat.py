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
        client = ChatClient(mock_credential)
        
        mock_response = MagicMock()
        mock_response.status_code = 201
        mock_response.json.return_value = {"id": "new-conv-123"}
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.return_value = mock_response
            
            result = await client.create_conversation()
            
            assert result == "new-conv-123"
            mock_req.assert_called_once()

    @pytest.mark.asyncio
    async def test_create_conversation_failure(self, mock_credential):
        """Should raise ChatApiError on failure."""
        client = ChatClient(mock_credential)
        
        mock_response = MagicMock()
        mock_response.status_code = 400
        mock_response.text = "Bad request"
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.return_value = mock_response
            
            with pytest.raises(ChatApiError):
                await client.create_conversation()
