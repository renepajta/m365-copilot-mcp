"""Base utilities for M365 Copilot API clients.

Provides common functionality for all M365 Copilot API clients:
- Request ID generation for log correlation
- Query truncation for GDPR compliance
- Response formatting
"""

from __future__ import annotations

import logging
import os
import secrets
from dataclasses import dataclass
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

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
