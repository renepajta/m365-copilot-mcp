"""Search API client for M365 Copilot.

Semantic document search in OneDrive.
Uses official Microsoft SDK (microsoft-agents-m365copilot-beta).

Endpoint: POST /copilot/search
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from microsoft_agents_m365copilot_beta import AgentsM365CopilotBetaServiceClient
from microsoft_agents_m365copilot_beta.generated.copilot.search.search_post_request_body import (
    SearchPostRequestBody,
)

from m365_copilot.clients.base import (
    gen_request_id,
    truncate_query,
)

if TYPE_CHECKING:
    from azure.core.credentials import TokenCredential

logger = logging.getLogger(__name__)

# Search API timeout
SEARCH_TIMEOUT = 60


@dataclass
class SearchResult:
    """A document from the Search API."""

    name: str
    url: str
    preview: str = ""
    file_type: str | None = None
    size: int | None = None
    last_modified: str | None = None
    author: str | None = None
    path: str | None = None

    def to_markdown(self) -> str:
        """Format result as markdown."""
        lines = []

        # Title with link
        lines.append(f"**[{self.name}]({self.url})**")

        # Metadata line
        meta = []
        if self.file_type:
            meta.append(self.file_type.upper())
        if self.size:
            size_kb = self.size / 1024
            if size_kb > 1024:
                meta.append(f"{size_kb / 1024:.1f} MB")
            else:
                meta.append(f"{size_kb:.0f} KB")
        if self.author:
            meta.append(f"by {self.author}")
        if meta:
            lines.append(f"*{' | '.join(meta)}*")

        # Preview
        if self.preview:
            lines.append(f"\n{self.preview}")

        lines.append("")
        return "\n".join(lines)


@dataclass
class SearchResponse:
    """Response from M365 Copilot Search API."""

    results: list[SearchResult] = field(default_factory=list)
    total_results: int = 0

    def to_markdown(self) -> str:
        """Format all results as markdown."""
        if not self.results:
            return "No documents found matching your query."

        lines = [f"Found {len(self.results)} documents:\n"]

        for i, result in enumerate(self.results, 1):
            lines.append(f"### {i}. {result.to_markdown()}")

        return "\n".join(lines)


class SearchClient:
    """Client for M365 Copilot Search API using official Microsoft SDK."""

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        self.credential = credential
        self.timeout = timeout or SEARCH_TIMEOUT
        
        # Create SDK client with correct beta API configuration
        from m365_copilot.auth import create_sdk_client
        self._sdk_client = create_sdk_client(credential)

    async def search(
        self,
        query: str,
        *,
        path_filter: str | None = None,
        page_size: int = 25,
        request_id: str | None = None,
    ) -> SearchResponse:
        """Search for documents in OneDrive.

        Args:
            query: Natural language search query.
            path_filter: Optional path filter (e.g., '/Documents/Projects').
            page_size: Number of results (1-100).

        Returns:
            SearchResponse with document results.
        """
        request_id = request_id or gen_request_id()

        logger.info(
            "[%s] Search: %s (path=%s, size=%d)",
            request_id,
            truncate_query(query),
            path_filter or "all",
            page_size,
        )

        # Build request body using SDK model
        request_body = SearchPostRequestBody()
        request_body.query = query
        request_body.page_size = min(max(1, page_size), 100)  # Clamp to 1-100

        # Add path filter via additional_data if provided
        if path_filter:
            request_body.additional_data["filter"] = {"path": path_filter}

        try:
            # Call SDK search endpoint
            result = await self._sdk_client.copilot.search.post(request_body)
            
            if result is None:
                return SearchResponse(results=[], total_results=0)
            
            results = self._parse_results_from_sdk(result)
            
            logger.info(
                "[%s] Found %d documents",
                request_id,
                len(results),
            )

            return SearchResponse(
                results=results,
                total_results=result.total_count or len(results),
            )
            
        except Exception as e:
            logger.error(
                "[%s] Search failed: %s",
                request_id,
                str(e),
            )
            raise SearchApiError(f"Search failed: {e}")

    def _parse_results_from_sdk(self, result: Any) -> list[SearchResult]:
        """Parse results from SDK response."""
        results = []

        if not hasattr(result, 'search_hits') or result.search_hits is None:
            return results

        for hit in result.search_hits:
            # Extract metadata from hit
            name = "Untitled"
            url = hit.web_url or ""
            preview = hit.preview or ""
            file_type = None
            size = None
            last_modified = None
            author = None
            path = None
            
            if hit.resource_type:
                file_type = str(hit.resource_type.value) if hasattr(hit.resource_type, 'value') else str(hit.resource_type)
            
            # Extract from resource_metadata if available
            if hit.resource_metadata and hasattr(hit.resource_metadata, 'additional_data'):
                metadata = hit.resource_metadata.additional_data
                name = metadata.get('name', name)
                size = metadata.get('size')
                last_modified = metadata.get('lastModifiedDateTime')
                author = metadata.get('lastModifiedBy', {}).get('user', {}).get('displayName') if isinstance(metadata.get('lastModifiedBy'), dict) else None
                path = metadata.get('parentReference', {}).get('path') if isinstance(metadata.get('parentReference'), dict) else None

            result_item = SearchResult(
                name=name,
                url=url,
                preview=preview,
                file_type=file_type,
                size=size,
                last_modified=last_modified,
                author=author,
                path=path,
            )
            results.append(result_item)

        return results


class SearchApiError(Exception):
    """Error from M365 Copilot Search API."""

    pass
