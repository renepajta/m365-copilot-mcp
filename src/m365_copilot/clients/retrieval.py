"""Retrieval API client for M365 Copilot.

Retrieves text chunks from SharePoint/OneDrive for RAG scenarios.
Uses official Microsoft SDK (microsoft-agents-m365copilot-beta).

Endpoint: POST /copilot/retrieval
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal

from microsoft_agents_m365copilot_beta import AgentsM365CopilotBetaServiceClient
from microsoft_agents_m365copilot_beta.generated.models.retrieval_data_source import (
    RetrievalDataSource,
)
from microsoft_agents_m365copilot_beta.generated.copilot.retrieval.retrieval_post_request_body import (
    RetrievalPostRequestBody,
)

from m365_copilot.clients.base import (
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


class RetrievalClient:
    """Client for M365 Copilot Retrieval API using official Microsoft SDK."""

    # Data source type mapping to SDK enum values
    DATA_SOURCE_MAP = {
        "sharepoint": RetrievalDataSource.SharePoint,
        "onedrive": RetrievalDataSource.OneDriveBusiness,
        "connectors": RetrievalDataSource.ExternalItem,
    }

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        self.credential = credential
        self.timeout = timeout or RETRIEVAL_TIMEOUT
        
        # Create SDK client with correct beta API configuration
        from m365_copilot.auth import create_sdk_client
        self._sdk_client = create_sdk_client(credential)

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

        logger.info(
            "[%s] Retrieve: %s (source=%s, max=%d)",
            request_id,
            truncate_query(query),
            data_source,
            max_results,
        )

        # Build request body using SDK models
        request_body = RetrievalPostRequestBody()
        request_body.query_string = query
        request_body.data_source = self.DATA_SOURCE_MAP.get(
            data_source, RetrievalDataSource.SharePoint
        )
        request_body.maximum_number_of_results = min(max(1, max_results), 25)
        
        if filter_expression:
            request_body.filter_expression = filter_expression

        try:
            # Call SDK retrieval endpoint
            result = await self._sdk_client.copilot.retrieval.post(request_body)
            
            if result is None:
                return RetrievalResponse(chunks=[], total_results=0)
            
            chunks = self._parse_chunks_from_sdk(result)
            
            logger.info(
                "[%s] Retrieved %d chunks",
                request_id,
                len(chunks),
            )

            return RetrievalResponse(
                chunks=chunks,
                total_results=len(chunks),
            )
            
        except Exception as e:
            logger.error(
                "[%s] Retrieval failed: %s",
                request_id,
                str(e),
            )
            raise RetrievalApiError(f"Retrieval failed: {e}")

    def _parse_chunks_from_sdk(self, result: Any) -> list[TextChunk]:
        """Parse chunks from SDK response."""
        chunks = []

        # SDK returns RetrievalResponse with retrieval_hits
        if not hasattr(result, 'retrieval_hits') or result.retrieval_hits is None:
            return chunks
            
        for hit in result.retrieval_hits:
            web_url = hit.web_url or ""
            
            # Get metadata from hit
            metadata = hit.resource_metadata
            title = None
            last_modified = None
            if metadata and hasattr(metadata, 'additional_data'):
                title = metadata.additional_data.get('title')
                last_modified = metadata.additional_data.get('lastModifiedDateTime')
            
            resource_type = None
            if hit.resource_type:
                resource_type = str(hit.resource_type.value) if hasattr(hit.resource_type, 'value') else str(hit.resource_type)
            
            # Each hit can have multiple extracts
            if hit.extracts:
                for extract in hit.extracts:
                    text_content = ""
                    relevance_score = 0.0
                    
                    if hasattr(extract, 'text'):
                        text_content = extract.text or ""
                    if hasattr(extract, 'relevance_score'):
                        relevance_score = extract.relevance_score or 0.0
                    
                    chunk = TextChunk(
                        content=text_content,
                        relevance_score=relevance_score,
                        source_url=web_url,
                        source_title=title,
                        file_type=resource_type,
                        last_modified=last_modified,
                    )
                    chunks.append(chunk)

        # Sort by relevance score descending
        chunks.sort(key=lambda c: c.relevance_score, reverse=True)

        return chunks


class RetrievalApiError(Exception):
    """Error from M365 Copilot Retrieval API."""

    pass
