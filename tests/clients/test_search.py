"""Tests for Search API client."""

import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from m365_copilot.clients.search import (
    SearchClient,
    SearchResponse,
    SearchResult,
    SearchApiError,
)


class TestSearchResult:
    """Tests for SearchResult dataclass."""

    def test_to_markdown(self):
        """Should format result as markdown."""
        result = SearchResult(
            name="Test Document.docx",
            url="https://example.com/doc",
            preview="This is a preview of the document...",
            file_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            size=102400,  # 100 KB
            author="John Doe",
        )
        
        markdown = result.to_markdown()
        assert "Test Document.docx" in markdown
        assert "https://example.com/doc" in markdown
        assert "John Doe" in markdown

    def test_to_markdown_large_file(self):
        """Should format large files in MB."""
        result = SearchResult(
            name="BigFile.pptx",
            url="https://example.com/big",
            size=5242880,  # 5 MB
        )
        
        markdown = result.to_markdown()
        assert "MB" in markdown


class TestSearchResponse:
    """Tests for SearchResponse dataclass."""

    def test_to_markdown_empty(self):
        """Should handle empty results."""
        response = SearchResponse(results=[], total_results=0)
        markdown = response.to_markdown()
        assert "No documents found" in markdown

    def test_to_markdown_with_results(self):
        """Should format multiple results."""
        response = SearchResponse(
            results=[
                SearchResult(name="Doc1.pdf", url="https://example.com/1"),
                SearchResult(name="Doc2.docx", url="https://example.com/2"),
            ],
            total_results=2,
        )
        
        markdown = response.to_markdown()
        assert "Found 2 documents" in markdown
        assert "Doc1.pdf" in markdown
        assert "Doc2.docx" in markdown


class TestSearchClient:
    """Tests for SearchClient."""

    @pytest.fixture
    def mock_credential(self):
        """Create mock credential."""
        cred = MagicMock()
        cred.get_token.return_value = MagicMock(token="test-token")
        return cred

    @pytest.fixture
    def mock_sdk_client(self):
        """Create mock SDK client."""
        mock_client = MagicMock()
        mock_client.copilot = MagicMock()
        mock_client.copilot.search = MagicMock()
        mock_client.copilot.search.post = AsyncMock()
        return mock_client

    @pytest.mark.asyncio
    async def test_search_success(self, mock_credential, mock_sdk_client):
        """Should search and parse results."""
        # Create mock SDK response
        mock_hit = MagicMock()
        mock_hit.web_url = "https://example.com/report"
        mock_hit.preview = "Quarterly report summary..."
        mock_hit.resource_type = None
        mock_hit.resource_metadata = MagicMock()
        mock_hit.resource_metadata.additional_data = {
            "name": "Report.docx",
            "size": 50000,
        }
        
        mock_response = MagicMock()
        mock_response.search_hits = [mock_hit]
        mock_response.total_count = 1
        
        mock_sdk_client.copilot.search.post.return_value = mock_response
        
        with patch(
            "m365_copilot.auth.create_sdk_client",
            return_value=mock_sdk_client,
        ):
            client = SearchClient(mock_credential)
            result = await client.search("quarterly report")
            
            assert len(result.results) == 1
            assert result.results[0].name == "Report.docx"
            assert result.total_results == 1

    @pytest.mark.asyncio
    async def test_search_with_path_filter(self, mock_credential, mock_sdk_client):
        """Should include path filter in request."""
        mock_response = MagicMock()
        mock_response.search_hits = []
        mock_response.total_count = 0
        
        mock_sdk_client.copilot.search.post.return_value = mock_response
        
        with patch(
            "m365_copilot.auth.create_sdk_client",
            return_value=mock_sdk_client,
        ):
            client = SearchClient(mock_credential)
            await client.search(
                "test query",
                path_filter="/Documents/Projects",
            )
            
            # Verify SDK was called
            mock_sdk_client.copilot.search.post.assert_called_once()
            call_args = mock_sdk_client.copilot.search.post.call_args
            request_body = call_args[0][0]
            assert request_body.query == "test query"

    @pytest.mark.asyncio
    async def test_search_failure(self, mock_credential, mock_sdk_client):
        """Should raise SearchApiError on failure."""
        mock_sdk_client.copilot.search.post.side_effect = Exception("API error")
        
        with patch(
            "m365_copilot.auth.create_sdk_client",
            return_value=mock_sdk_client,
        ):
            client = SearchClient(mock_credential)
            
            with pytest.raises(SearchApiError):
                await client.search("test query")

    @pytest.mark.asyncio
    async def test_search_page_size_clamped(self, mock_credential, mock_sdk_client):
        """Should clamp page_size to valid range."""
        mock_response = MagicMock()
        mock_response.search_hits = []
        mock_response.total_count = 0
        
        mock_sdk_client.copilot.search.post.return_value = mock_response
        
        with patch(
            "m365_copilot.auth.create_sdk_client",
            return_value=mock_sdk_client,
        ):
            client = SearchClient(mock_credential)
            
            # Test with value above max
            await client.search("test", page_size=500)
            
            call_args = mock_sdk_client.copilot.search.post.call_args
            request_body = call_args[0][0]
            assert request_body.page_size <= 100

    @pytest.mark.asyncio
    async def test_search_empty_response(self, mock_credential, mock_sdk_client):
        """Should handle empty/null response."""
        mock_sdk_client.copilot.search.post.return_value = None
        
        with patch(
            "m365_copilot.auth.create_sdk_client",
            return_value=mock_sdk_client,
        ):
            client = SearchClient(mock_credential)
            result = await client.search("test query")
            
            assert len(result.results) == 0
            assert result.total_results == 0
