"""Base Graph client with shared utilities.

Provides common functionality for all M365 Copilot API clients:
- Request ID generation for log correlation
- Query truncation for GDPR compliance
- Retry logic with exponential backoff
- Response formatting
"""

from __future__ import annotations

import logging
import os
import secrets
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Any

import httpx
from msgraph import GraphServiceClient
from msgraph_core import GraphClientFactory

if TYPE_CHECKING:
    from azure.core.credentials import TokenCredential

logger = logging.getLogger(__name__)

# Retry configuration
MAX_RETRIES = 3
RETRY_DELAYS = [5, 15, 30]  # seconds

# Default timeout (can be overridden per-client)
DEFAULT_TIMEOUT = 60


def gen_request_id() -> str:
    """Generate a 6-character hex request ID for log correlation."""
    return secrets.token_hex(3)


def truncate_query(query: str, max_length: int = 50) -> str:
    """Truncate query for logging (GDPR compliance)."""
    if len(query) <= max_length:
        return query
    return query[:max_length] + "..."


def get_timeout() -> int:
    """Get timeout from environment or default."""
    timeout_str = os.getenv("M365_COPILOT_TIMEOUT")
    if timeout_str:
        try:
            return int(timeout_str)
        except ValueError:
            logger.warning("Invalid M365_COPILOT_TIMEOUT value: %s", timeout_str)
    return DEFAULT_TIMEOUT


@dataclass
class Attribution:
    """Source attribution from M365 Copilot response."""

    type: str  # 'annotation', 'citation', 'grounding'
    text: str
    url: str | None = None
    title: str | None = None


@dataclass
class UsageStats:
    """Usage statistics for a request."""

    request_id: str
    started_at: datetime
    completed_at: datetime | None = None
    latency_ms: int | None = None

    def complete(self) -> None:
        """Mark request as complete and calculate latency."""
        self.completed_at = datetime.now(timezone.utc)
        if self.started_at:
            delta = self.completed_at - self.started_at
            self.latency_ms = int(delta.total_seconds() * 1000)


class GraphClient:
    """Base client for Microsoft Graph API calls.

    Uses msgraph-sdk for standard REST calls.
    Subclasses may use httpx directly for streaming (SSE) endpoints.
    """

    BETA_BASE_URL = "https://graph.microsoft.com/beta"
    V1_BASE_URL = "https://graph.microsoft.com/v1.0"

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        """Initialize Graph client.

        Args:
            credential: Azure credential for authentication.
            timeout: Request timeout in seconds.
        """
        self.credential = credential
        self.timeout = timeout or get_timeout()

        # Create msgraph SDK client
        self._graph_client = GraphServiceClient(
            credential,
            scopes=["https://graph.microsoft.com/.default"],
        )

        # Create httpx client for custom requests (SSE, beta endpoints)
        self._http_client: httpx.AsyncClient | None = None

    async def _get_http_client(self) -> httpx.AsyncClient:
        """Get or create async HTTP client."""
        if self._http_client is None:
            self._http_client = httpx.AsyncClient(
                timeout=httpx.Timeout(self.timeout),
                follow_redirects=True,
            )
        return self._http_client

    async def _get_access_token(self) -> str:
        """Get access token from credential."""
        from m365_copilot.auth import GRAPH_SCOPES

        token = self.credential.get_token(*GRAPH_SCOPES)
        return token.token

    async def _make_request(
        self,
        method: str,
        url: str,
        *,
        json: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
        request_id: str | None = None,
    ) -> httpx.Response:
        """Make an authenticated HTTP request.

        Args:
            method: HTTP method.
            url: Full URL to request.
            json: JSON body (optional).
            headers: Additional headers (optional).
            request_id: Request ID for logging (optional).

        Returns:
            httpx.Response object.
        """
        request_id = request_id or gen_request_id()
        token = await self._get_access_token()

        all_headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "X-Request-Id": request_id,
        }
        if headers:
            all_headers.update(headers)

        client = await self._get_http_client()

        logger.debug("[%s] %s %s", request_id, method, url)

        response = await client.request(
            method,
            url,
            json=json,
            headers=all_headers,
        )

        logger.debug("[%s] Response: %d", request_id, response.status_code)
        return response

    async def close(self) -> None:
        """Close HTTP client."""
        if self._http_client:
            await self._http_client.aclose()
            self._http_client = None


def format_citations(attributions: list[Attribution]) -> str:
    """Format attributions as markdown citations section.

    Returns markdown like:
    ---
    **Sources:**
    [^1^]: [Title](url)
    [^2^]: [Title](url)
    """
    if not attributions:
        return ""

    lines = ["\n---", "**Sources:**"]
    for i, attr in enumerate(attributions, 1):
        if attr.url:
            title = attr.title or attr.text or f"Source {i}"
            lines.append(f"[^{i}^]: [{title}]({attr.url})")
        elif attr.text:
            lines.append(f"[^{i}^]: {attr.text}")

    return "\n".join(lines)


def format_sensitivity_label(label: str | None) -> str:
    """Format sensitivity label as footer warning.

    Returns markdown like:
    ---
    ⚠️ Sensitivity: Confidential
    """
    if not label:
        return ""
    return f"\n---\n⚠️ Sensitivity: {label}"
