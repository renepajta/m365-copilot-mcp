"""Tests for Meeting Insights API client."""

import pytest
from unittest.mock import AsyncMock, MagicMock, patch
from datetime import datetime, timezone

from m365_copilot.clients.meetings import (
    MeetingsClient,
    MeetingInsight,
    MeetingNote,
    ActionItem,
    MentionEvent,
    MeetingSummary,
    MeetingsApiError,
)


class TestMeetingNote:
    """Tests for MeetingNote dataclass."""

    def test_creation(self):
        """Should create note with subpoints."""
        note = MeetingNote(
            title="Discussion Topic",
            text="Main discussion points...",
            subpoints=[
                MeetingNote(title="Point 1", text="Detail 1"),
                MeetingNote(title="Point 2", text="Detail 2"),
            ],
        )
        
        assert note.title == "Discussion Topic"
        assert len(note.subpoints) == 2


class TestMeetingInsight:
    """Tests for MeetingInsight dataclass."""

    def test_to_markdown_empty(self):
        """Should handle insight with no data."""
        insight = MeetingInsight(meeting_id="meeting-123")
        markdown = insight.to_markdown()
        assert "No insights available" in markdown

    def test_to_markdown_with_notes(self):
        """Should format notes section."""
        insight = MeetingInsight(
            meeting_id="meeting-123",
            meeting_title="Sprint Planning",
            notes=[
                MeetingNote(
                    title="Sprint Goals",
                    text="Complete feature X by end of sprint",
                )
            ],
        )
        
        markdown = insight.to_markdown()
        assert "Sprint Planning" in markdown
        assert "Summary" in markdown
        assert "Sprint Goals" in markdown

    def test_to_markdown_with_action_items(self):
        """Should format action items section."""
        insight = MeetingInsight(
            meeting_id="meeting-123",
            action_items=[
                ActionItem(
                    title="Review PR",
                    text="Review the authentication PR",
                    owner="Alice",
                ),
            ],
        )
        
        markdown = insight.to_markdown()
        assert "Action Items" in markdown
        assert "Review PR" in markdown
        assert "@Alice" in markdown

    def test_to_markdown_with_mentions(self):
        """Should format mentions section."""
        insight = MeetingInsight(
            meeting_id="meeting-123",
            mentions=[
                MentionEvent(
                    timestamp="2026-01-10T14:30:00Z",
                    text="We need input from you on this",
                    speaker="Bob",
                ),
            ],
        )
        
        markdown = insight.to_markdown()
        assert "You Were Mentioned" in markdown
        assert "Bob" in markdown


class TestMeetingSummary:
    """Tests for MeetingSummary dataclass."""

    def test_to_markdown(self):
        """Should format as list item."""
        summary = MeetingSummary(
            meeting_id="meeting-123",
            title="Team Standup",
            start_time="2026-01-10T09:00:00Z",
        )
        
        markdown = summary.to_markdown()
        assert "Team Standup" in markdown
        assert "meeting-123" in markdown


class TestMeetingsClient:
    """Tests for MeetingsClient."""

    @pytest.fixture
    def mock_credential(self):
        """Create mock credential."""
        cred = MagicMock()
        cred.get_token.return_value = MagicMock(token="test-token")
        return cred

    @pytest.fixture
    def mock_sdk_client(self):
        """Create mock SDK client."""
        mock_client = MagicMock()
        
        # Mock copilot.users hierarchy
        mock_client.copilot = MagicMock()
        mock_client.copilot.users = MagicMock()
        
        # Create mock for by_ai_user_id chain
        mock_user = MagicMock()
        mock_client.copilot.users.by_ai_user_id = MagicMock(return_value=mock_user)
        
        # Mock online_meetings
        mock_user.online_meetings = MagicMock()
        mock_user.online_meetings.get = AsyncMock()
        
        # Mock by_ai_online_meeting_id chain
        mock_meeting = MagicMock()
        mock_user.online_meetings.by_ai_online_meeting_id = MagicMock(return_value=mock_meeting)
        
        # Mock ai_insights
        mock_meeting.ai_insights = MagicMock()
        mock_meeting.ai_insights.get = AsyncMock()
        
        return mock_client

    @pytest.mark.asyncio
    async def test_list_meetings_success(self, mock_credential, mock_sdk_client):
        """Should list meetings."""
        # Mock meetings response
        mock_meeting_item = MagicMock()
        mock_meeting_item.id = "meeting-1"
        mock_meeting_item.subject = "Team Meeting"
        mock_meeting_item.start_date_time = datetime(2026, 1, 10, 9, 0, 0, tzinfo=timezone.utc)
        mock_meeting_item.join_web_url = "https://teams.microsoft.com/..."
        
        mock_meetings_response = MagicMock()
        mock_meetings_response.value = [mock_meeting_item]
        
        mock_user_obj = mock_sdk_client.copilot.users.by_ai_user_id.return_value
        mock_user_obj.online_meetings.get.return_value = mock_meetings_response
        
        with patch("m365_copilot.auth.create_sdk_client", return_value=mock_sdk_client):
            with patch.object(MeetingsClient, "_get_current_user_id", new_callable=AsyncMock) as mock_get_user:
                mock_get_user.return_value = "user-123"
                
                client = MeetingsClient(mock_credential)
                result = await client.list_meetings()
                
                assert len(result) == 1
                assert result[0].meeting_id == "meeting-1"
                assert result[0].title == "Team Meeting"

    @pytest.mark.asyncio
    async def test_get_insights_not_found(self, mock_credential, mock_sdk_client):
        """Should return placeholder when insights not available (empty response)."""
        # Mock insights response (empty value list)
        mock_insights_response = MagicMock()
        mock_insights_response.value = []
        
        mock_user_obj = mock_sdk_client.copilot.users.by_ai_user_id.return_value
        mock_meeting_obj = mock_user_obj.online_meetings.by_ai_online_meeting_id.return_value
        mock_meeting_obj.ai_insights.get.return_value = mock_insights_response
        
        with patch("m365_copilot.auth.create_sdk_client", return_value=mock_sdk_client):
            with patch.object(MeetingsClient, "_get_current_user_id", new_callable=AsyncMock) as mock_get_user:
                mock_get_user.return_value = "user-123"
                
                client = MeetingsClient(mock_credential)
                result = await client.get_insights("meeting-123")
                
                assert result.meeting_id == "meeting-123"
                # Empty response returns placeholder note
                assert len(result.notes) == 1
                assert "not yet available" in result.notes[0].text.lower()

    @pytest.mark.asyncio
    async def test_get_insights_404_error(self, mock_credential, mock_sdk_client):
        """Should return placeholder when 404 error."""
        # Mock 404 error
        mock_user_obj = mock_sdk_client.copilot.users.by_ai_user_id.return_value
        mock_meeting_obj = mock_user_obj.online_meetings.by_ai_online_meeting_id.return_value
        mock_meeting_obj.ai_insights.get.side_effect = Exception("404 Not Found")
        
        with patch("m365_copilot.auth.create_sdk_client", return_value=mock_sdk_client):
            with patch.object(MeetingsClient, "_get_current_user_id", new_callable=AsyncMock) as mock_get_user:
                mock_get_user.return_value = "user-123"
                
                client = MeetingsClient(mock_credential)
                result = await client.get_insights("meeting-123")
                
                assert result.meeting_id == "meeting-123"
                assert len(result.notes) == 1
                assert "not yet available" in result.notes[0].text.lower()

    @pytest.mark.asyncio
    async def test_get_insights_success(self, mock_credential, mock_sdk_client):
        """Should parse full insights response."""
        # Create mock insight data
        mock_note = MagicMock()
        mock_note.title = "Overview"
        mock_note.text = "Discussed roadmap"
        mock_note.subpoints = []
        
        mock_action = MagicMock()
        mock_action.title = "Draft spec"
        mock_action.text = "Write spec doc"
        mock_action.owner_display_name = "Alice"
        
        mock_speaker = MagicMock()
        mock_speaker.display_name = "Bob"
        
        mock_mention = MagicMock()
        mock_mention.event_date_time = "2026-01-10T10:00:00Z"
        mock_mention.transcript_utterance = "Can you review this?"
        mock_mention.speaker = mock_speaker
        
        mock_viewpoint = MagicMock()
        mock_viewpoint.mention_events = [mock_mention]
        
        mock_insight = MagicMock()
        mock_insight.subject = "Planning Meeting"
        mock_insight.meeting_notes = [mock_note]
        mock_insight.action_items = [mock_action]
        mock_insight.viewpoint = mock_viewpoint
        
        mock_insights_response = MagicMock()
        mock_insights_response.value = [mock_insight]
        
        mock_user_obj = mock_sdk_client.copilot.users.by_ai_user_id.return_value
        mock_meeting_obj = mock_user_obj.online_meetings.by_ai_online_meeting_id.return_value
        mock_meeting_obj.ai_insights.get.return_value = mock_insights_response
        
        with patch("m365_copilot.auth.create_sdk_client", return_value=mock_sdk_client):
            with patch.object(MeetingsClient, "_get_current_user_id", new_callable=AsyncMock) as mock_get_user:
                mock_get_user.return_value = "user-123"
                
                client = MeetingsClient(mock_credential)
                result = await client.get_insights("meeting-123")
                
                assert result.meeting_title == "Planning Meeting"
                assert len(result.notes) == 1
                assert len(result.action_items) == 1
                assert len(result.mentions) == 1

    def test_extract_meeting_id_from_url(self, mock_credential, mock_sdk_client):
        """Should extract meeting ID from Teams URL."""
        with patch(
            "m365_copilot.auth.create_sdk_client",
            return_value=mock_sdk_client,
        ):
            client = MeetingsClient(mock_credential)
            
            url = "https://teams.microsoft.com/l/meetup-join/ABC123XYZ/0"
            result = client._extract_meeting_id(url)
            
            assert result == "ABC123XYZ"

    def test_extract_meeting_id_invalid_url(self, mock_credential, mock_sdk_client):
        """Should return empty string for invalid URL."""
        with patch(
            "m365_copilot.auth.create_sdk_client",
            return_value=mock_sdk_client,
        ):
            client = MeetingsClient(mock_credential)
            
            result = client._extract_meeting_id("https://example.com")
            assert result == ""
