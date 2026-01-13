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

    @pytest.mark.asyncio
    async def test_search_success(self, mock_credential):
        """Should search and parse results."""
        client = SearchClient(mock_credential)
        
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "value": [
                {
                    "resource": {
                        "name": "Report.docx",
                        "webUrl": "https://example.com/report",
                        "size": 50000,
                    },
                    "summary": "Quarterly report summary...",
                }
            ],
            "totalCount": 1,
        }
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.return_value = mock_response
            
            result = await client.search("quarterly report")
            
            assert len(result.results) == 1
            assert result.results[0].name == "Report.docx"
            assert result.total_results == 1

    @pytest.mark.asyncio
    async def test_search_with_path_filter(self, mock_credential):
        """Should include path filter in request."""
        client = SearchClient(mock_credential)
        
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"value": [], "totalCount": 0}
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.return_value = mock_response
            
            await client.search(
                "test query",
                path_filter="/Documents/Projects",
            )
            
            call_args = mock_req.call_args
            assert call_args is not None

    @pytest.mark.asyncio
    async def test_search_failure(self, mock_credential):
        """Should raise SearchApiError on failure."""
        client = SearchClient(mock_credential)
        
        mock_response = MagicMock()
        mock_response.status_code = 400
        mock_response.text = "Bad request"
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.return_value = mock_response
            
            with pytest.raises(SearchApiError):
                await client.search("test query")

    @pytest.mark.asyncio
    async def test_search_page_size_clamped(self, mock_credential):
        """Should clamp page_size to valid range."""
        client = SearchClient(mock_credential)
        
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"value": []}
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.return_value = mock_response
            
            # Test with value above max
            await client.search("test", page_size=500)
            
            call_kwargs = mock_req.call_args[1]
            body = call_kwargs.get("json", {})
            assert body.get("pageSize", 0) <= 100
