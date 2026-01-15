"""Tests for Retrieval API client."""

import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from m365_copilot.clients.retrieval import (
    RetrievalClient,
    RetrievalResponse,
    TextChunk,
    RetrievalApiError,
)
from microsoft_agents_m365copilot_beta.generated.models.retrieval_data_source import (
    RetrievalDataSource,
)


class TestTextChunk:
    """Tests for TextChunk dataclass."""

    def test_to_markdown(self):
        """Should format chunk as markdown."""
        chunk = TextChunk(
            content="This is the content.",
            relevance_score=0.95,
            source_url="https://example.com/doc",
            source_title="Test Document",
        )
        
        result = chunk.to_markdown()
        assert "Test Document" in result
        assert "0.95" in result
        assert "This is the content." in result


class TestRetrievalResponse:
    """Tests for RetrievalResponse dataclass."""

    def test_to_markdown_empty(self):
        """Should handle empty results."""
        response = RetrievalResponse(chunks=[], total_results=0)
        result = response.to_markdown()
        assert "No relevant content found" in result

    def test_to_markdown_with_chunks(self):
        """Should format multiple chunks."""
        response = RetrievalResponse(
            chunks=[
                TextChunk(content="First chunk", relevance_score=0.9),
                TextChunk(content="Second chunk", relevance_score=0.8),
            ],
            total_results=2,
        )
        
        result = response.to_markdown()
        assert "Found 2 relevant chunks" in result
        assert "First chunk" in result
        assert "Second chunk" in result


class TestRetrievalClient:
    """Tests for RetrievalClient."""

    @pytest.fixture
    def mock_credential(self):
        """Create mock credential."""
        cred = MagicMock()
        cred.get_token.return_value = MagicMock(token="test-token")
        return cred

    @pytest.mark.asyncio
    async def test_retrieve_success(self, mock_credential):
        """Should retrieve and parse chunks."""
        with patch(
            "m365_copilot.auth.create_sdk_client"
        ) as mock_sdk_class:
            mock_sdk = MagicMock()
            mock_sdk_class.return_value = mock_sdk
            
            # Mock SDK response
            mock_extract = MagicMock()
            mock_extract.text = "Test content"
            mock_extract.relevance_score = 0.85
            
            mock_hit = MagicMock()
            mock_hit.web_url = "https://example.com/doc"
            mock_hit.resource_metadata = MagicMock()
            mock_hit.resource_metadata.additional_data = {"title": "Test Doc"}
            mock_hit.resource_type = None
            mock_hit.extracts = [mock_extract]
            
            mock_result = MagicMock()
            mock_result.retrieval_hits = [mock_hit]
            
            mock_sdk.copilot.retrieval.post = AsyncMock(return_value=mock_result)
            
            client = RetrievalClient(mock_credential)
            result = await client.retrieve("test query")
            
            assert len(result.chunks) == 1
            assert result.chunks[0].content == "Test content"
            assert result.chunks[0].relevance_score == 0.85

    @pytest.mark.asyncio
    async def test_retrieve_with_filter(self, mock_credential):
        """Should include filter in request."""
        with patch(
            "m365_copilot.auth.create_sdk_client"
        ) as mock_sdk_class:
            mock_sdk = MagicMock()
            mock_sdk_class.return_value = mock_sdk
            
            mock_result = MagicMock()
            mock_result.retrieval_hits = []
            
            mock_sdk.copilot.retrieval.post = AsyncMock(return_value=mock_result)
            
            client = RetrievalClient(mock_credential)
            await client.retrieve(
                "test query",
                filter_expression="FileType:pdf",
            )
            
            # Check that filter was included in request body
            call_args = mock_sdk.copilot.retrieval.post.call_args
            assert call_args is not None
            request_body = call_args[0][0]
            assert request_body.filter_expression == "FileType:pdf"

    @pytest.mark.asyncio
    async def test_retrieve_failure(self, mock_credential):
        """Should raise RetrievalApiError on failure."""
        with patch(
            "m365_copilot.auth.create_sdk_client"
        ) as mock_sdk_class:
            mock_sdk = MagicMock()
            mock_sdk_class.return_value = mock_sdk
            mock_sdk.copilot.retrieval.post = AsyncMock(
                side_effect=Exception("API error")
            )
            
            client = RetrievalClient(mock_credential)
            
            with pytest.raises(RetrievalApiError):
                await client.retrieve("test query")

    def test_data_source_mapping(self, mock_credential):
        """Should map data source types correctly."""
        with patch(
            "m365_copilot.auth.create_sdk_client"
        ):
            client = RetrievalClient(mock_credential)
            
            assert client.DATA_SOURCE_MAP["sharepoint"] == RetrievalDataSource.SharePoint
            assert client.DATA_SOURCE_MAP["onedrive"] == RetrievalDataSource.OneDriveBusiness
            assert client.DATA_SOURCE_MAP["connectors"] == RetrievalDataSource.ExternalItem
