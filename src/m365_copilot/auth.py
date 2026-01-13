"""Authentication for Microsoft Graph API.

Implements delegated authentication flows:
- Interactive browser (default for local dev)
- Device code flow (fallback for headless environments)
- Token caching for subsequent runs
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import TYPE_CHECKING

from azure.identity import (
    ChainedTokenCredential,
    DeviceCodeCredential,
    InteractiveBrowserCredential,
    TokenCachePersistenceOptions,
)

if TYPE_CHECKING:
    from azure.core.credentials import TokenCredential

logger = logging.getLogger(__name__)

# Microsoft Graph scopes required for M365 Copilot APIs
# Chat API requires ALL of these simultaneously
GRAPH_SCOPES = [
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/People.Read.All",
    "https://graph.microsoft.com/OnlineMeetingTranscript.Read.All",
    "https://graph.microsoft.com/Chat.Read",
    "https://graph.microsoft.com/ChannelMessage.Read.All",
    "https://graph.microsoft.com/ExternalItem.Read.All",
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/OnlineMeeting.Read",
]

# Default cache directory
DEFAULT_CACHE_DIR = Path.home() / ".m365-copilot-mcp"


def get_cache_dir() -> Path:
    """Get the token cache directory from env or default."""
    cache_dir = os.getenv("M365_COPILOT_CACHE_DIR")
    if cache_dir:
        return Path(cache_dir).expanduser()
    return DEFAULT_CACHE_DIR


def get_credential(
    client_id: str | None = None,
    tenant_id: str | None = None,
    *,
    allow_browser: bool = True,
) -> TokenCredential:
    """Get a chained credential for Microsoft Graph authentication.

    Priority order:
    1. Interactive browser (if allowed and available)
    2. Device code flow (fallback)

    Args:
        client_id: Azure AD app client ID. Defaults to AZURE_CLIENT_ID env var.
        tenant_id: Azure AD tenant ID. Defaults to AZURE_TENANT_ID env var.
        allow_browser: Whether to try interactive browser auth first.

    Returns:
        A TokenCredential that can be used with Microsoft Graph SDK.

    Raises:
        ValueError: If client_id or tenant_id is not provided and not in env.
    """
    client_id = client_id or os.getenv("AZURE_CLIENT_ID")
    tenant_id = tenant_id or os.getenv("AZURE_TENANT_ID")

    if not client_id:
        raise ValueError(
            "AZURE_CLIENT_ID is required. Set it in environment or pass client_id."
        )
    if not tenant_id:
        raise ValueError(
            "AZURE_TENANT_ID is required. Set it in environment or pass tenant_id."
        )

    # Ensure cache directory exists
    cache_dir = get_cache_dir()
    cache_dir.mkdir(parents=True, exist_ok=True)

    # Configure token cache persistence
    cache_options = TokenCachePersistenceOptions(
        name="m365-copilot-mcp",
        allow_unencrypted_storage=True,  # Required for WSL/Linux without keyring
    )

    credentials: list[TokenCredential] = []

    if allow_browser:
        # Interactive browser - best for local dev with GUI
        try:
            browser_cred = InteractiveBrowserCredential(
                client_id=client_id,
                tenant_id=tenant_id,
                redirect_uri="http://localhost:8400",
                cache_persistence_options=cache_options,
            )
            credentials.append(browser_cred)
            logger.debug("Added InteractiveBrowserCredential to chain")
        except Exception as e:
            logger.warning("Could not create browser credential: %s", e)

    # Device code flow - fallback for headless/SSH
    device_cred = DeviceCodeCredential(
        client_id=client_id,
        tenant_id=tenant_id,
        cache_persistence_options=cache_options,
        prompt_callback=_device_code_prompt,
    )
    credentials.append(device_cred)
    logger.debug("Added DeviceCodeCredential to chain")

    return ChainedTokenCredential(*credentials)


def _device_code_prompt(
    verification_uri: str,
    user_code: str,
    expires_on: object,
) -> None:
    """Callback to display device code instructions."""
    logger.info(
        "To sign in, visit %s and enter code: %s",
        verification_uri,
        user_code,
    )
    print(f"\n🔐 To authenticate, visit: {verification_uri}")
    print(f"   Enter code: {user_code}\n")


async def get_access_token(
    credential: TokenCredential,
    scopes: list[str] | None = None,
) -> str:
    """Get an access token for Microsoft Graph.

    Args:
        credential: The TokenCredential to use.
        scopes: OAuth scopes to request. Defaults to GRAPH_SCOPES.

    Returns:
        Access token string.
    """
    scopes = scopes or GRAPH_SCOPES

    # azure-identity credentials have get_token, not async by default
    # but we wrap it for consistency in async context
    token = credential.get_token(*scopes)
    return token.token


def clear_token_cache() -> None:
    """Clear the local token cache (for troubleshooting)."""
    cache_dir = get_cache_dir()
    if cache_dir.exists():
        import shutil
        shutil.rmtree(cache_dir)
        logger.info("Cleared token cache at %s", cache_dir)
