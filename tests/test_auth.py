"""Tests for authentication module."""

import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path

from m365_copilot.auth import (
    get_cache_dir,
    get_credential,
    GRAPH_SCOPES,
    DEFAULT_CACHE_DIR,
    _load_auth_record,
    _save_auth_record,
    _get_auth_record_path,
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

    @patch("m365_copilot.auth._load_auth_record", return_value=None)
    @patch("m365_copilot.auth.ChainedTokenCredential")
    @patch("m365_copilot.auth.DeviceCodeCredential")
    @patch("m365_copilot.auth.InteractiveBrowserCredential")
    def test_creates_chained_credential(
        self, mock_browser, mock_device, mock_chained, mock_load_record
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

    @patch("m365_copilot.auth._load_auth_record", return_value=None)
    @patch("m365_copilot.auth.ChainedTokenCredential")
    @patch("m365_copilot.auth.DeviceCodeCredential")
    @patch("m365_copilot.auth.SharedTokenCacheCredential")
    def test_no_browser_when_disabled(self, mock_shared, mock_device, mock_chained, mock_load_record):
        """Should skip browser credential when allow_browser=False."""
        with patch.dict(
            "os.environ",
            {"AZURE_CLIENT_ID": "client123", "AZURE_TENANT_ID": "tenant123"},
        ):
            with patch("pathlib.Path.mkdir"):
                get_credential(allow_browser=False)

        mock_device.assert_called_once()
        # Chained should have shared cache + device credential (no browser)
        call_args = mock_chained.call_args[0]
        assert len(call_args) == 2  # SharedTokenCacheCredential + DeviceCodeCredential


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


class TestAuthRecord:
    """Tests for authentication record persistence."""

    def test_auth_record_path_uses_cache_dir(self):
        """Auth record path should be under cache directory."""
        with patch.dict("os.environ", {}, clear=True):
            path = _get_auth_record_path()
            assert path.parent == DEFAULT_CACHE_DIR
            assert path.name == "auth_record.json"

    def test_load_auth_record_returns_none_when_missing(self, tmp_path):
        """Should return None when auth record doesn't exist."""
        with patch("m365_copilot.auth.get_cache_dir", return_value=tmp_path):
            result = _load_auth_record()
            assert result is None

    @patch("m365_copilot.auth.AuthenticationRecord")
    def test_uses_saved_auth_record_when_available(self, mock_auth_record):
        """Should use saved auth record for silent authentication."""
        mock_record = MagicMock()
        mock_record.username = "user@example.com"

        with patch.dict(
            "os.environ",
            {"AZURE_CLIENT_ID": "client123", "AZURE_TENANT_ID": "tenant123"},
        ):
            with patch("m365_copilot.auth._load_auth_record", return_value=mock_record):
                with patch("m365_copilot.auth.ChainedTokenCredential") as mock_chained:
                    with patch("m365_copilot.auth.InteractiveBrowserCredential") as mock_browser:
                        with patch("m365_copilot.auth.SharedTokenCacheCredential"):
                            with patch("m365_copilot.auth.DeviceCodeCredential"):
                                with patch("pathlib.Path.mkdir"):
                                    get_credential()

        # Should create InteractiveBrowserCredential twice:
        # 1. Silent credential with auth record
        # 2. Interactive fallback
        assert mock_browser.call_count == 2

        # First call should have auth record and disable_automatic_authentication
        first_call = mock_browser.call_args_list[0]
        assert first_call.kwargs.get("authentication_record") == mock_record
        assert first_call.kwargs.get("disable_automatic_authentication") is True
