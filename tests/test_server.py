"""Tests for the MCP server module."""

import pytest
from unittest.mock import AsyncMock, MagicMock, patch


class TestServerEndpoints:
    """Tests for HTTP endpoints."""

    @pytest.mark.asyncio
    async def test_root_info(self):
        """Should return service info."""
        from m365_copilot.server import root_info
        
        response = await root_info(None)
        data = response.body.decode()
        
        assert "m365-copilot-mcp" in data
        assert "healthy" in data or "running" in data

    @pytest.mark.asyncio
    async def test_health_check(self):
        """Should return healthy status."""
        from m365_copilot.server import health_check
        
        response = await health_check(None)
        data = response.body.decode()
        
        assert "healthy" in data


class TestToolHelpers:
    """Tests for tool helper functions."""

    def test_gen_request_id(self):
        """Should generate 6-char hex ID."""
        from m365_copilot.clients.base import gen_request_id
        
        rid = gen_request_id()
        assert len(rid) == 6
        assert all(c in "0123456789abcdef" for c in rid)

    def test_truncate_query_short(self):
        """Should not truncate short queries."""
        from m365_copilot.clients.base import truncate_query
        
        result = truncate_query("short query", 50)
        assert result == "short query"

    def test_truncate_query_long(self):
        """Should truncate long queries."""
        from m365_copilot.clients.base import truncate_query
        
        long_query = "a" * 100
        result = truncate_query(long_query, 50)
        assert len(result) == 53  # 50 + "..."
        assert result.endswith("...")


class TestClientInitialization:
    """Tests for client lazy initialization."""

    @patch("m365_copilot.server.get_credential")
    def test_get_chat_client(self, mock_cred):
        """Should create chat client on first call."""
        from m365_copilot import server
        server._credential = None
        server._chat_client = None
        
        mock_cred.return_value = MagicMock()
        
        client = server._get_chat_client()
        assert client is not None
        
        # Second call should return same instance
        client2 = server._get_chat_client()
        assert client is client2

    @patch("m365_copilot.server.get_credential")
    def test_get_retrieval_client(self, mock_cred):
        """Should create retrieval client on first call."""
        from m365_copilot import server
        server._credential = None
        server._retrieval_client = None
        
        mock_cred.return_value = MagicMock()
        
        client = server._get_retrieval_client()
        assert client is not None

    @patch("m365_copilot.server.get_credential")
    def test_get_search_client(self, mock_cred):
        """Should create search client on first call."""
        from m365_copilot import server
        server._credential = None
        server._search_client = None
        
        mock_cred.return_value = MagicMock()
        
        client = server._get_search_client()
        assert client is not None

    @patch("m365_copilot.server.get_credential")
    def test_get_meetings_client(self, mock_cred):
        """Should create meetings client on first call."""
        from m365_copilot import server
        server._credential = None
        server._meetings_client = None
        
        mock_cred.return_value = MagicMock()
        
        client = server._get_meetings_client()
        assert client is not None
