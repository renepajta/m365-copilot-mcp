"""Search API client for M365 Copilot.

Semantic document search in OneDrive.

Endpoint: POST /beta/copilot/search
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from m365_copilot.clients.base import (
    GraphClient,
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


class SearchClient(GraphClient):
    """Client for M365 Copilot Search API."""

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        super().__init__(credential, timeout=timeout or SEARCH_TIMEOUT)

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
        url = f"{self.BETA_BASE_URL}/copilot/search"

        logger.info(
            "[%s] Search: %s (path=%s, size=%d)",
            request_id,
            truncate_query(query),
            path_filter or "all",
            page_size,
        )

        # Build request body
        body: dict[str, Any] = {
            "query": query,
            "pageSize": min(max(1, page_size), 100),  # Clamp to 1-100
        }

        # Add path filter if provided
        if path_filter:
            body["filter"] = {
                "path": path_filter,
            }

        response = await self._make_request(
            "POST",
            url,
            json=body,
            request_id=request_id,
        )

        if response.status_code != 200:
            logger.error(
                "[%s] Search failed: %d %s",
                request_id,
                response.status_code,
                response.text,
            )
            raise SearchApiError(
                f"Search failed: {response.status_code} - {response.text}"
            )

        data = response.json()
        results = self._parse_results(data)

        logger.info(
            "[%s] Found %d documents",
            request_id,
            len(results),
        )

        return SearchResponse(
            results=results,
            total_results=data.get("totalCount", len(results)),
        )

    def _parse_results(self, data: dict[str, Any]) -> list[SearchResult]:
        """Parse results from API response."""
        results = []

        for item in data.get("value", []):
            # Extract resource data
            resource = item.get("resource", {})

            result = SearchResult(
                name=resource.get("name", "Untitled"),
                url=resource.get("webUrl", ""),
                preview=item.get("summary", ""),
                file_type=resource.get("file", {}).get("mimeType"),
                size=resource.get("size"),
                last_modified=resource.get("lastModifiedDateTime"),
                author=resource.get("lastModifiedBy", {})
                .get("user", {})
                .get("displayName"),
                path=resource.get("parentReference", {}).get("path"),
            )
            results.append(result)

        return results


class SearchApiError(Exception):
    """Error from M365 Copilot Search API."""

    pass
