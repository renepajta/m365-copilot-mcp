"""Tests for authentication module."""

import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path

from m365_copilot.auth import (
    get_cache_dir,
    get_credential,
    GRAPH_SCOPES,
    DEFAULT_CACHE_DIR,
)


class TestGetCacheDir:
    """Tests for get_cache_dir function."""

    def test_default_cache_dir(self):
        """Should return default cache dir when env not set."""
        with patch.dict("os.environ", {}, clear=True):
            result = get_cache_dir()
            assert result == DEFAULT_CACHE_DIR

    def test_env_override(self):
        """Should use M365_COPILOT_CACHE_DIR from environment."""
        with patch.dict("os.environ", {"M365_COPILOT_CACHE_DIR": "/custom/cache"}):
            result = get_cache_dir()
            assert result == Path("/custom/cache")

    def test_tilde_expansion(self):
        """Should expand ~ in path."""
        with patch.dict("os.environ", {"M365_COPILOT_CACHE_DIR": "~/mycache"}):
            result = get_cache_dir()
            assert "~" not in str(result)


class TestGetCredential:
    """Tests for get_credential function."""

    def test_missing_client_id(self):
        """Should raise ValueError when client_id is missing."""
        with patch.dict("os.environ", {"AZURE_TENANT_ID": "tenant123"}, clear=True):
            with pytest.raises(ValueError, match="AZURE_CLIENT_ID is required"):
                get_credential()

    def test_missing_tenant_id(self):
        """Should raise ValueError when tenant_id is missing."""
        with patch.dict("os.environ", {"AZURE_CLIENT_ID": "client123"}, clear=True):
            with pytest.raises(ValueError, match="AZURE_TENANT_ID is required"):
                get_credential()

    @patch("m365_copilot.auth.ChainedTokenCredential")
    @patch("m365_copilot.auth.DeviceCodeCredential")
    @patch("m365_copilot.auth.InteractiveBrowserCredential")
    def test_creates_chained_credential(
        self, mock_browser, mock_device, mock_chained
    ):
        """Should create chained credential with browser and device code."""
        with patch.dict(
            "os.environ",
            {"AZURE_CLIENT_ID": "client123", "AZURE_TENANT_ID": "tenant123"},
        ):
            with patch("pathlib.Path.mkdir"):
                get_credential()

        mock_browser.assert_called_once()
        mock_device.assert_called_once()
        mock_chained.assert_called_once()

    @patch("m365_copilot.auth.ChainedTokenCredential")
    @patch("m365_copilot.auth.DeviceCodeCredential")
    def test_no_browser_when_disabled(self, mock_device, mock_chained):
        """Should skip browser credential when allow_browser=False."""
        with patch.dict(
            "os.environ",
            {"AZURE_CLIENT_ID": "client123", "AZURE_TENANT_ID": "tenant123"},
        ):
            with patch("pathlib.Path.mkdir"):
                get_credential(allow_browser=False)

        mock_device.assert_called_once()
        # Chained should only have device credential
        call_args = mock_chained.call_args[0]
        assert len(call_args) == 1


class TestGraphScopes:
    """Tests for Graph API scopes."""

    def test_required_scopes_present(self):
        """Should include all required Chat API scopes."""
        required = [
            "Sites.Read.All",
            "Mail.Read",
            "Files.Read.All",
            "OnlineMeetingTranscript.Read.All",
        ]
        for scope in required:
            assert any(scope in s for s in GRAPH_SCOPES)

    def test_scopes_are_fully_qualified(self):
        """All scopes should be fully qualified with graph.microsoft.com."""
        for scope in GRAPH_SCOPES:
            assert scope.startswith("https://graph.microsoft.com/")
