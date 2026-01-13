"""Meeting Insights API client for M365 Copilot.

Gets AI-generated summaries and action items from Teams meetings.

Endpoints:
- GET /v1.0/copilot/users/{userId}/onlineMeetings/{meetingId}/aiInsights
- GET /v1.0/copilot/users/{userId}/onlineMeetings/{meetingId}/aiInsights/{insightId}
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import TYPE_CHECKING, Any

from m365_copilot.clients.base import (
    GraphClient,
    gen_request_id,
)

if TYPE_CHECKING:
    from azure.core.credentials import TokenCredential

logger = logging.getLogger(__name__)

# Meeting Insights API timeout
MEETINGS_TIMEOUT = 30


@dataclass
class MeetingNote:
    """A note/summary point from a meeting."""

    title: str
    text: str
    subpoints: list[MeetingNote] = field(default_factory=list)


@dataclass
class ActionItem:
    """An action item from a meeting."""

    title: str
    text: str
    owner: str | None = None
    due_date: str | None = None


@dataclass
class MentionEvent:
    """A mention of the user in a meeting."""

    timestamp: str
    text: str
    speaker: str


@dataclass
class MeetingInsight:
    """AI-generated insights from a Teams meeting."""

    meeting_id: str
    meeting_title: str | None = None
    meeting_date: str | None = None
    notes: list[MeetingNote] = field(default_factory=list)
    action_items: list[ActionItem] = field(default_factory=list)
    mentions: list[MentionEvent] = field(default_factory=list)

    def to_markdown(self) -> str:
        """Format insight as markdown."""
        lines = []

        # Meeting header
        if self.meeting_title:
            lines.append(f"# {self.meeting_title}")
        if self.meeting_date:
            lines.append(f"*{self.meeting_date}*\n")

        # Meeting notes
        if self.notes:
            lines.append("## Summary")
            for note in self.notes:
                lines.append(f"### {note.title}")
                lines.append(note.text)
                for sub in note.subpoints:
                    lines.append(f"- **{sub.title}**: {sub.text}")
                lines.append("")

        # Action items
        if self.action_items:
            lines.append("## Action Items")
            for item in self.action_items:
                owner_str = f" (@{item.owner})" if item.owner else ""
                lines.append(f"- [ ] **{item.title}**{owner_str}")
                lines.append(f"  {item.text}")
            lines.append("")

        # Mentions
        if self.mentions:
            lines.append("## You Were Mentioned")
            for mention in self.mentions:
                lines.append(f"- *{mention.speaker}* at {mention.timestamp}:")
                lines.append(f"  > {mention.text}")
            lines.append("")

        return "\n".join(lines) if lines else "No insights available for this meeting."


@dataclass
class MeetingSummary:
    """Brief summary of a meeting for listing."""

    meeting_id: str
    title: str
    start_time: str
    join_url: str | None = None

    def to_markdown(self) -> str:
        """Format as markdown list item."""
        return f"- **{self.title}** ({self.start_time})\n  ID: `{self.meeting_id}`"


class MeetingsClient(GraphClient):
    """Client for M365 Copilot Meeting Insights API."""

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        super().__init__(credential, timeout=timeout or MEETINGS_TIMEOUT)

    async def list_meetings(
        self,
        *,
        since: datetime | None = None,
        request_id: str | None = None,
    ) -> list[MeetingSummary]:
        """List recent meetings with available insights.

        Args:
            since: Only include meetings after this datetime.
                   Defaults to 7 days ago.

        Returns:
            List of meeting summaries.
        """
        request_id = request_id or gen_request_id()

        # Default to last 7 days
        if since is None:
            since = datetime.now(timezone.utc) - timedelta(days=7)

        # Get current user ID
        user_id = await self._get_current_user_id(request_id)

        # List meetings with calendar events that have online meeting info
        url = (
            f"{self.V1_BASE_URL}/users/{user_id}/onlineMeetings"
            f"?$filter=startDateTime ge {since.isoformat()}"
            f"&$orderby=startDateTime desc"
            f"&$top=50"
        )

        logger.info("[%s] Listing meetings since %s", request_id, since.date())

        response = await self._make_request(
            "GET",
            url,
            request_id=request_id,
        )

        if response.status_code != 200:
            logger.error(
                "[%s] List meetings failed: %d %s",
                request_id,
                response.status_code,
                response.text,
            )
            raise MeetingsApiError(
                f"Failed to list meetings: {response.status_code}"
            )

        data = response.json()
        meetings = []

        for item in data.get("value", []):
            meeting = MeetingSummary(
                meeting_id=item.get("id", ""),
                title=item.get("subject", "Untitled Meeting"),
                start_time=item.get("startDateTime", ""),
                join_url=item.get("joinWebUrl"),
            )
            meetings.append(meeting)

        logger.info("[%s] Found %d meetings", request_id, len(meetings))
        return meetings

    async def get_insights(
        self,
        meeting_id: str,
        *,
        join_url: str | None = None,
        request_id: str | None = None,
    ) -> MeetingInsight:
        """Get AI insights for a specific meeting.

        Args:
            meeting_id: Teams meeting ID.
            join_url: Optional join URL (can be used instead of meeting_id).

        Returns:
            MeetingInsight with notes, action items, and mentions.
        """
        request_id = request_id or gen_request_id()

        # If join URL provided, extract meeting ID
        if join_url and not meeting_id:
            meeting_id = self._extract_meeting_id(join_url)

        if not meeting_id:
            raise MeetingsApiError("Either meeting_id or join_url is required")

        user_id = await self._get_current_user_id(request_id)

        url = (
            f"{self.V1_BASE_URL}/copilot/users/{user_id}"
            f"/onlineMeetings/{meeting_id}/aiInsights"
        )

        logger.info("[%s] Getting insights for meeting %s", request_id, meeting_id)

        response = await self._make_request(
            "GET",
            url,
            request_id=request_id,
        )

        if response.status_code == 404:
            return MeetingInsight(
                meeting_id=meeting_id,
                notes=[
                    MeetingNote(
                        title="No Insights Available",
                        text="AI insights are not yet available for this meeting. "
                        "Insights typically become available ~4 hours after the meeting ends. "
                        "Ensure transcription was enabled during the meeting.",
                    )
                ],
            )

        if response.status_code != 200:
            logger.error(
                "[%s] Get insights failed: %d %s",
                request_id,
                response.status_code,
                response.text,
            )
            raise MeetingsApiError(
                f"Failed to get meeting insights: {response.status_code}"
            )

        data = response.json()
        insight = self._parse_insight(meeting_id, data)

        logger.info(
            "[%s] Got insights: %d notes, %d actions, %d mentions",
            request_id,
            len(insight.notes),
            len(insight.action_items),
            len(insight.mentions),
        )

        return insight

    async def _get_current_user_id(self, request_id: str) -> str:
        """Get the current user's ID from Graph."""
        url = f"{self.V1_BASE_URL}/me"

        response = await self._make_request(
            "GET",
            url,
            request_id=request_id,
        )

        if response.status_code != 200:
            raise MeetingsApiError("Failed to get current user info")

        data = response.json()
        return data.get("id", "")

    def _extract_meeting_id(self, join_url: str) -> str:
        """Extract meeting ID from Teams join URL."""
        # Teams URLs contain encoded meeting ID
        # Example: https://teams.microsoft.com/l/meetup-join/...
        match = re.search(r"meetup-join/([^/]+)/", join_url)
        if match:
            return match.group(1)
        return ""

    def _parse_insight(self, meeting_id: str, data: dict[str, Any]) -> MeetingInsight:
        """Parse insight from API response."""
        # Get the first insight (usually only one per meeting)
        insights_list = data.get("value", [])
        if not insights_list:
            return MeetingInsight(meeting_id=meeting_id)

        insight_data = insights_list[0]

        # Parse meeting notes
        notes = []
        for note_data in insight_data.get("meetingNotes", []):
            note = MeetingNote(
                title=note_data.get("title", ""),
                text=note_data.get("text", ""),
                subpoints=[
                    MeetingNote(
                        title=sub.get("title", ""),
                        text=sub.get("text", ""),
                    )
                    for sub in note_data.get("subpoints", [])
                ],
            )
            notes.append(note)

        # Parse action items
        action_items = []
        for item_data in insight_data.get("actionItems", []):
            item = ActionItem(
                title=item_data.get("title", ""),
                text=item_data.get("text", ""),
                owner=item_data.get("ownerDisplayName"),
            )
            action_items.append(item)

        # Parse mention events (from viewpoint)
        mentions = []
        viewpoint = insight_data.get("viewpoint", {})
        for mention_data in viewpoint.get("mentionEvents", []):
            mention = MentionEvent(
                timestamp=mention_data.get("eventDateTime", ""),
                text=mention_data.get("transcriptUtterance", ""),
                speaker=mention_data.get("speaker", {}).get("displayName", "Unknown"),
            )
            mentions.append(mention)

        return MeetingInsight(
            meeting_id=meeting_id,
            meeting_title=insight_data.get("subject"),
            meeting_date=insight_data.get("startDateTime"),
            notes=notes,
            action_items=action_items,
            mentions=mentions,
        )


class MeetingsApiError(Exception):
    """Error from M365 Copilot Meeting Insights API."""

    pass
