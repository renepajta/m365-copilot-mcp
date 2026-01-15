"""Meeting Insights API client for M365 Copilot.

Gets AI-generated summaries and action items from Teams meetings.
Uses official Microsoft SDK (microsoft-agents-m365copilot-beta).

Endpoints:
- GET /copilot/users/{userId}/onlineMeetings
- GET /copilot/users/{userId}/onlineMeetings/{meetingId}/aiInsights
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import TYPE_CHECKING, Any

import httpx
from microsoft_agents_m365copilot_beta import AgentsM365CopilotBetaServiceClient

from m365_copilot.clients.base import gen_request_id

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


class MeetingsClient:
    """Client for M365 Copilot Meeting Insights API using official Microsoft SDK."""

    def __init__(
        self,
        credential: TokenCredential,
        *,
        timeout: int | None = None,
    ) -> None:
        self.credential = credential
        self.timeout = timeout or MEETINGS_TIMEOUT
        
        # Create SDK client with correct beta API configuration
        from m365_copilot.auth import create_sdk_client
        self._sdk_client = create_sdk_client(credential)

    async def list_meetings(
        self,
        *,
        since: datetime | None = None,
        request_id: str | None = None,
    ) -> list[MeetingSummary]:
        """List recent meetings with available insights.

        Note: The standard Graph /me/onlineMeetings endpoint requires a filter.
        This method uses the /me/calendar/events endpoint with $filter for meetings
        since the Copilot-specific endpoint may not be available in all tenants.

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

        logger.info("[%s] Listing meetings since %s", request_id, since.date())

        # First try the Copilot-specific endpoint
        try:
            return await self._list_meetings_copilot(since, request_id)
        except MeetingsApiError as e:
            if "NotFound" in str(e) or "not supported" in str(e).lower():
                logger.info("[%s] Copilot meetings endpoint not available, using calendar events", request_id)
            else:
                raise

        # Fall back to calendar events (which shows Teams meetings)
        return await self._list_meetings_calendar(since, request_id)

    async def _list_meetings_copilot(
        self,
        since: datetime,
        request_id: str,
    ) -> list[MeetingSummary]:
        """List meetings using Copilot-specific endpoint."""
        user_id = await self._get_current_user_id(request_id)

        try:
            # Use SDK to get online meetings
            result = await self._sdk_client.copilot.users.by_ai_user_id(
                user_id
            ).online_meetings.get()
            
            meetings = []
            
            if result and result.value:
                for item in result.value:
                    # Filter by start time
                    if hasattr(item, 'start_date_time') and item.start_date_time:
                        meeting_time = item.start_date_time
                        if isinstance(meeting_time, datetime):
                            if meeting_time < since:
                                continue
                    
                    meeting = MeetingSummary(
                        meeting_id=item.id or "",
                        title=getattr(item, 'subject', None) or "Untitled Meeting",
                        start_time=str(item.start_date_time) if hasattr(item, 'start_date_time') and item.start_date_time else "",
                        join_url=getattr(item, 'join_web_url', None),
                    )
                    meetings.append(meeting)

            logger.info("[%s] Found %d meetings via Copilot endpoint", request_id, len(meetings))
            return meetings
            
        except Exception as e:
            raise MeetingsApiError(f"Failed to list meetings: {e}")

    async def _list_meetings_calendar(
        self,
        since: datetime,
        request_id: str,
    ) -> list[MeetingSummary]:
        """List meetings using calendar events endpoint (fallback).
        
        Uses /me/calendar/calendarView to get meetings with Teams join URLs.
        """
        from m365_copilot.auth import GRAPH_SCOPES
        
        try:
            token = self.credential.get_token(*GRAPH_SCOPES)
            
            # Use calendarView with date range
            since_str = since.strftime('%Y-%m-%dT%H:%M:%SZ')
            until = datetime.now(timezone.utc) + timedelta(days=30)  # Include upcoming
            until_str = until.strftime('%Y-%m-%dT%H:%M:%SZ')
            
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    "https://graph.microsoft.com/v1.0/me/calendar/calendarView",
                    params={
                        "startDateTime": since_str,
                        "endDateTime": until_str,
                        "$filter": "isOnlineMeeting eq true",
                        "$select": "id,subject,start,end,onlineMeeting,isOnlineMeeting",
                        "$orderby": "start/dateTime desc",
                        "$top": "50",
                    },
                    headers={"Authorization": f"Bearer {token.token}"},
                    timeout=30.0,
                )
                response.raise_for_status()
                data = response.json()
            
            meetings = []
            for event in data.get("value", []):
                # Extract meeting info
                online_meeting = event.get("onlineMeeting", {}) or {}
                join_url = online_meeting.get("joinUrl")
                
                # Try to extract meeting ID from join URL
                meeting_id = ""
                if join_url:
                    match = re.search(r"19:meeting_([^/]+)", join_url)
                    if match:
                        meeting_id = match.group(0)
                
                start_info = event.get("start", {})
                start_time = start_info.get("dateTime", "")
                
                meeting = MeetingSummary(
                    meeting_id=meeting_id or event.get("id", ""),
                    title=event.get("subject") or "Untitled Meeting",
                    start_time=start_time,
                    join_url=join_url,
                )
                meetings.append(meeting)
            
            logger.info("[%s] Found %d meetings via calendar", request_id, len(meetings))
            return meetings
            
        except httpx.HTTPStatusError as e:
            raise MeetingsApiError(f"Failed to list meetings: HTTP {e.response.status_code}")
        except Exception as e:
            raise MeetingsApiError(f"Failed to list meetings: {e}")

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

        logger.info("[%s] Getting insights for meeting %s", request_id, meeting_id)

        try:
            # Use SDK to get AI insights
            # The SDK endpoint: copilot/users/{userId}/onlineMeetings/{meetingId}/aiInsights
            result = await self._sdk_client.copilot.users.by_ai_user_id(
                user_id
            ).online_meetings.by_ai_online_meeting_id(
                meeting_id
            ).ai_insights.get()
            
            if result is None or (hasattr(result, 'value') and not result.value):
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
            
            insight = self._parse_insight_from_sdk(meeting_id, result)

            logger.info(
                "[%s] Got insights: %d notes, %d actions, %d mentions",
                request_id,
                len(insight.notes),
                len(insight.action_items),
                len(insight.mentions),
            )

            return insight
            
        except Exception as e:
            # Check if it's a 404-like error
            error_str = str(e).lower()
            if "404" in error_str or "not found" in error_str:
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
            
            logger.error(
                "[%s] Get insights failed: %s",
                request_id,
                str(e),
            )
            raise MeetingsApiError(f"Failed to get meeting insights: {e}")

    async def _get_current_user_id(self, request_id: str) -> str:
        """Get the current user's ID from Graph /me endpoint."""
        from m365_copilot.auth import GRAPH_SCOPES
        
        try:
            # Get access token
            token = self.credential.get_token(*GRAPH_SCOPES)
            
            # Call /me endpoint with raw HTTP (M365 Copilot SDK doesn't include /me)
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    "https://graph.microsoft.com/v1.0/me",
                    headers={"Authorization": f"Bearer {token.token}"},
                    timeout=10.0,
                )
                response.raise_for_status()
                data = response.json()
                
                user_id = data.get("id")
                if not user_id:
                    raise MeetingsApiError("Failed to get current user info: no ID returned")
                return user_id
                
        except httpx.HTTPStatusError as e:
            raise MeetingsApiError(f"Failed to get current user info: HTTP {e.response.status_code}")
        except Exception as e:
            raise MeetingsApiError(f"Failed to get current user info: {e}")

    def _extract_meeting_id(self, join_url: str) -> str:
        """Extract meeting ID from Teams join URL."""
        # Teams URLs contain encoded meeting ID
        # Example: https://teams.microsoft.com/l/meetup-join/...
        match = re.search(r"meetup-join/([^/]+)/", join_url)
        if match:
            return match.group(1)
        return ""

    def _parse_insight_from_sdk(self, meeting_id: str, data: Any) -> MeetingInsight:
        """Parse insight from SDK response."""
        # Get the first insight (usually only one per meeting)
        insights_list = data.value if hasattr(data, 'value') and data.value else []
        if not insights_list:
            return MeetingInsight(meeting_id=meeting_id)

        insight_data = insights_list[0]

        # Parse meeting notes
        notes = []
        meeting_notes_list = getattr(insight_data, 'meeting_notes', None) or []
        for note_data in meeting_notes_list:
            subpoints_list = getattr(note_data, 'subpoints', None) or []
            note = MeetingNote(
                title=getattr(note_data, 'title', '') or '',
                text=getattr(note_data, 'text', '') or '',
                subpoints=[
                    MeetingNote(
                        title=getattr(sub, 'title', '') or '',
                        text=getattr(sub, 'text', '') or '',
                    )
                    for sub in subpoints_list
                ],
            )
            notes.append(note)

        # Parse action items
        action_items = []
        action_items_list = getattr(insight_data, 'action_items', None) or []
        for item_data in action_items_list:
            item = ActionItem(
                title=getattr(item_data, 'title', '') or '',
                text=getattr(item_data, 'text', '') or '',
                owner=getattr(item_data, 'owner_display_name', None),
            )
            action_items.append(item)

        # Parse mention events (from viewpoint)
        mentions = []
        viewpoint = getattr(insight_data, 'viewpoint', None)
        if viewpoint:
            mention_events_list = getattr(viewpoint, 'mention_events', None) or []
            for mention_data in mention_events_list:
                speaker = getattr(mention_data, 'speaker', None)
                speaker_name = getattr(speaker, 'display_name', 'Unknown') if speaker else 'Unknown'
                mention = MentionEvent(
                    timestamp=str(getattr(mention_data, 'event_date_time', '')) or '',
                    text=getattr(mention_data, 'transcript_utterance', '') or '',
                    speaker=speaker_name,
                )
                mentions.append(mention)

        return MeetingInsight(
            meeting_id=meeting_id,
            meeting_title=getattr(insight_data, 'subject', None),
            meeting_date=str(getattr(insight_data, 'start_date_time', '')) if hasattr(insight_data, 'start_date_time') else None,
            notes=notes,
            action_items=action_items,
            mentions=mentions,
        )


class MeetingsApiError(Exception):
    """Error from M365 Copilot Meeting Insights API."""

    pass
