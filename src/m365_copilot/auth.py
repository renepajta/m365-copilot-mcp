"""Authentication for Microsoft Graph API.

Implements delegated authentication flows:
- SharedTokenCacheCredential (uses Azure CLI / shared MSAL cache)
- Interactive browser (fallback for fresh auth)
- Device code flow (for headless/stdio mode)

Token caching: Uses MSAL shared cache at ~/.azure/msal_token_cache.json
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import TYPE_CHECKING

from azure.identity import (
    AuthenticationRecord,
    ChainedTokenCredential,
    DeviceCodeCredential,
    InteractiveBrowserCredential,
    SharedTokenCacheCredential,
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
    "https://graph.microsoft.com/OnlineMeetings.Read",
]

# Default cache directory for our auth record
DEFAULT_CACHE_DIR = Path.home() / ".m365-copilot-mcp"
AUTH_RECORD_FILE = "auth_record.json"


def get_cache_dir() -> Path:
    """Get the token cache directory from env or default."""
    cache_dir = os.getenv("M365_COPILOT_CACHE_DIR")
    if cache_dir:
        return Path(cache_dir).expanduser()
    return DEFAULT_CACHE_DIR


def _get_auth_record_path() -> Path:
    """Get path to the authentication record file."""
    return get_cache_dir() / AUTH_RECORD_FILE


def _load_auth_record() -> AuthenticationRecord | None:
    """Load saved authentication record if it exists."""
    path = _get_auth_record_path()
    if path.exists():
        try:
            return AuthenticationRecord.deserialize(path.read_text())
        except Exception as e:
            logger.warning("Failed to load auth record: %s", e)
    return None


def _save_auth_record(record: AuthenticationRecord) -> None:
    """Save authentication record for future use."""
    path = _get_auth_record_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(record.serialize())
    logger.debug("Saved auth record to %s", path)


def get_credential(
    client_id: str | None = None,
    tenant_id: str | None = None,
    *,
    username: str | None = None,
    allow_browser: bool = True,
) -> TokenCredential:
    """Get a chained credential for Microsoft Graph authentication.

    Uses a priority chain:
    1. Saved AuthenticationRecord (silent token refresh)
    2. SharedTokenCacheCredential (Azure CLI cache) 
    3. InteractiveBrowserCredential (new login)
    4. DeviceCodeCredential (headless fallback)

    Args:
        client_id: Azure AD app client ID. Defaults to AZURE_CLIENT_ID env var.
        tenant_id: Azure AD tenant ID. Defaults to AZURE_TENANT_ID env var.
        username: Preferred username to use from cache (e.g., 'user@contoso.com').
                  Set via AZURE_USERNAME env var. Helps when multiple accounts cached.
        allow_browser: Whether to try interactive browser auth.

    Returns:
        A TokenCredential that can be used with Microsoft Graph SDK.

    Raises:
        ValueError: If client_id or tenant_id is not provided and not in env.
    """
    client_id = client_id or os.getenv("AZURE_CLIENT_ID")
    tenant_id = tenant_id or os.getenv("AZURE_TENANT_ID")
    username = username or os.getenv("AZURE_USERNAME")

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

    # Configure token cache persistence (uses MSAL shared cache)
    cache_options = TokenCachePersistenceOptions(
        name="m365-copilot-mcp",
        allow_unencrypted_storage=True,  # Required for WSL/Linux without keyring
    )

    credentials: list[TokenCredential] = []
    
    # 1. Try to use saved authentication record (enables silent refresh)
    auth_record = _load_auth_record()
    if auth_record:
        logger.debug("Found saved auth record for %s", auth_record.username)
        # Use the saved record with InteractiveBrowserCredential for silent auth
        silent_cred = InteractiveBrowserCredential(
            client_id=client_id,
            tenant_id=tenant_id,
            authentication_record=auth_record,
            cache_persistence_options=cache_options,
            disable_automatic_authentication=True,  # Don't prompt, just use cache
        )
        credentials.append(silent_cred)
    
    # 2. Try shared token cache (picks up Azure CLI login)
    try:
        shared_cred = SharedTokenCacheCredential(
            client_id=client_id,
            tenant_id=tenant_id,
            username=username,  # Filter to specific account if provided
        )
        credentials.append(shared_cred)
        logger.debug("Added SharedTokenCacheCredential (username=%s)", username or "any")
    except Exception as e:
        logger.debug("SharedTokenCacheCredential not available: %s", e)

    # 3. Interactive browser for new login
    if allow_browser:
        browser_cred = InteractiveBrowserCredential(
            client_id=client_id,
            tenant_id=tenant_id,
            redirect_uri="http://localhost:8400",
            cache_persistence_options=cache_options,
        )
        credentials.append(browser_cred)
        logger.debug("Added InteractiveBrowserCredential to chain")

    # 4. Device code flow as last resort
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


def authenticate_and_save(
    client_id: str | None = None,
    tenant_id: str | None = None,
) -> AuthenticationRecord:
    """Perform interactive authentication and save the record.
    
    Call this once to authenticate, then subsequent calls to get_credential()
    will silently use the cached tokens.
    
    Args:
        client_id: Azure AD app client ID.
        tenant_id: Azure AD tenant ID.
        
    Returns:
        The AuthenticationRecord that was saved.
    """
    client_id = client_id or os.getenv("AZURE_CLIENT_ID")
    tenant_id = tenant_id or os.getenv("AZURE_TENANT_ID")
    
    if not client_id or not tenant_id:
        raise ValueError("AZURE_CLIENT_ID and AZURE_TENANT_ID are required")
    
    cache_options = TokenCachePersistenceOptions(
        name="m365-copilot-mcp",
        allow_unencrypted_storage=True,
    )
    
    cred = InteractiveBrowserCredential(
        client_id=client_id,
        tenant_id=tenant_id,
        redirect_uri="http://localhost:8400",
        cache_persistence_options=cache_options,
    )
    
    # This will trigger interactive auth and return a record
    record = cred.authenticate(scopes=GRAPH_SCOPES)
    _save_auth_record(record)
    
    logger.info("Authenticated as %s and saved record", record.username)
    return record


def create_sdk_client(credential: TokenCredential) -> "AgentsM365CopilotBetaServiceClient":
    """Create M365 Copilot SDK client with correct beta API configuration.
    
    The official SDK has a bug where AgentsM365CopilotBetaRequestAdapter doesn't
    pass api_version=beta to the client factory, causing requests to hit v1.0
    endpoints instead of /beta endpoints. This function creates a properly
    configured client.
    
    Args:
        credential: Azure credential for authentication.
        
    Returns:
        Configured SDK client using /beta endpoints.
    """
    from kiota_authentication_azure.azure_identity_authentication_provider import (
        AzureIdentityAuthenticationProvider,
    )
    from microsoft_agents_m365copilot_beta._version import VERSION
    from microsoft_agents_m365copilot_beta.generated.base_agents_m365_copilot_beta_service_client import (
        BaseAgentsM365CopilotBetaServiceClient,
    )
    from microsoft_agents_m365copilot_core import (
        APIVersion,
        BaseMicrosoftAgentsM365CopilotRequestAdapter,
        MicrosoftAgentsM365CopilotClientFactory,
        MicrosoftAgentsM365CopilotTelemetryHandlerOption,
    )
    
    # Create auth provider
    auth_provider = AzureIdentityAuthenticationProvider(credential, scopes=GRAPH_SCOPES)
    
    # Create options with beta telemetry
    options = {
        MicrosoftAgentsM365CopilotTelemetryHandlerOption.get_key(): MicrosoftAgentsM365CopilotTelemetryHandlerOption(
            api_version=APIVersion.beta,
            sdk_version=VERSION,
        )
    }
    
    # Create HTTP client with CORRECT api_version=beta
    # (The SDK's AgentsM365CopilotBetaRequestAdapter has a bug - it doesn't pass this)
    http_client = MicrosoftAgentsM365CopilotClientFactory.create_with_default_middleware(
        api_version=APIVersion.beta,
        options=options,
    )
    
    # Create adapter with the properly configured HTTP client
    adapter = BaseMicrosoftAgentsM365CopilotRequestAdapter(
        auth_provider,
        http_client=http_client,
    )
    
    # Create and return the SDK client
    return BaseAgentsM365CopilotBetaServiceClient(adapter)
