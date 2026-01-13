"""Retrieval API client for M365 Copilot.

Retrieves text chunks from SharePoint/OneDrive for RAG scenarios.

Endpoint: POST /beta/copilot/retrieval
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal

from m365_copilot.clients.base import (
    GraphClient,
    gen_request_id,
    truncate_query,
)

if TYPE_CHECKING:
    from azure.core.credentials import TokenCredential

logger = logging.getLogger(__name__)

# Retrieval API timeout
RETRIEVAL_TIMEOUT = 90


@dataclass
class TextChunk:
    """A text chunk from the Retrieval API."""

    content: str
    relevance_score: float
    source_url: str | None = None
    source_title: str | None = None
    file_type: str | None = None
    last_modified: str | None = None

    def to_markdown(self) -> str:
        """Format chunk as markdown."""
        lines = []

        if self.source_title:
            lines.append(f"### {self.source_title}")
        if self.source_url:
            lines.append(f"*Source: [{self.source_url}]({self.source_url})*")
        if self.relevance_score:
            lines.append(f"*Relevance: {self.relevance_score:.2f}*")

        lines.append("")
        lines.append(self.content)
        lines.append("")

        return "\n".join(lines)


@dataclass
class RetrievalResponse:
    """Response from M365 Copilot Retrieval API."""

    chunks: list[TextChunk] = field(default_factory=list)
    total_results: int = 0

    def to_markdown(self) -> str:
        """Format all chunks as markdown."""
        if not self.chunks:
            return "No relevant content found."

        lines = [f"Found {len(self.chunks)} relevant chunks:\n"]

        for i, chunk in enumerate(self.chunks, 1):
            lines.append(f"---\n**[{i}]**\n")
            lines.append(chunk.to_markdown())

        return "\n".join(lines)


class RetrievalClient(GraphClient):
    """Client for M365 Copilot Retrieval API."""

    # Data source type mapping
    DATA_SOURCES = {
        "sharepoint": "microsoft365SharePoint",
        "onedrive": "microsoft365OneDrive",
        "connectors": "copilotConnectors",
    }

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        super().__init__(credential, timeout=timeout or RETRIEVAL_TIMEOUT)

    async def retrieve(
        self,
        query: str,
        *,
        data_source: Literal["sharepoint", "onedrive", "connectors"] = "sharepoint",
        filter_expression: str | None = None,
        max_results: int = 25,
        request_id: str | None = None,
    ) -> RetrievalResponse:
        """Retrieve text chunks from M365 for RAG.

        Args:
            query: Natural language search query.
            data_source: Where to search ('sharepoint', 'onedrive', 'connectors').
            filter_expression: Optional KQL filter expression.
            max_results: Maximum chunks to return (1-25).

        Returns:
            RetrievalResponse with text chunks.
        """
        request_id = request_id or gen_request_id()
        url = f"{self.BETA_BASE_URL}/copilot/retrieval"

        logger.info(
            "[%s] Retrieve: %s (source=%s, max=%d)",
            request_id,
            truncate_query(query),
            data_source,
            max_results,
        )

        # Build request body
        body: dict[str, Any] = {
            "query": query,
            "dataSource": {
                "type": self.DATA_SOURCES.get(data_source, data_source),
            },
            "maxResults": min(max(1, max_results), 25),  # Clamp to 1-25
        }

        # Add KQL filter if provided
        if filter_expression:
            body["dataSource"]["filterExpression"] = filter_expression

        response = await self._make_request(
            "POST",
            url,
            json=body,
            request_id=request_id,
        )

        if response.status_code != 200:
            logger.error(
                "[%s] Retrieval failed: %d %s",
                request_id,
                response.status_code,
                response.text,
            )
            raise RetrievalApiError(
                f"Retrieval failed: {response.status_code} - {response.text}"
            )

        data = response.json()
        chunks = self._parse_chunks(data)

        logger.info(
            "[%s] Retrieved %d chunks",
            request_id,
            len(chunks),
        )

        return RetrievalResponse(
            chunks=chunks,
            total_results=len(chunks),
        )

    def _parse_chunks(self, data: dict[str, Any]) -> list[TextChunk]:
        """Parse chunks from API response."""
        chunks = []

        for item in data.get("value", []):
            chunk = TextChunk(
                content=item.get("content", ""),
                relevance_score=item.get("relevanceScore", 0.0),
                source_url=item.get("webUrl"),
                source_title=item.get("name"),
                file_type=item.get("fileType"),
                last_modified=item.get("lastModifiedDateTime"),
            )
            chunks.append(chunk)

        # Sort by relevance score descending
        chunks.sort(key=lambda c: c.relevance_score, reverse=True)

        return chunks


class RetrievalApiError(Exception):
    """Error from M365 Copilot Retrieval API."""

    pass
