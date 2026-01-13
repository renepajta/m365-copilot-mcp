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

    @pytest.mark.asyncio
    async def test_list_meetings_success(self, mock_credential):
        """Should list meetings."""
        client = MeetingsClient(mock_credential)
        
        # Mock user ID request
        user_response = MagicMock()
        user_response.status_code = 200
        user_response.json.return_value = {"id": "user-123"}
        
        # Mock meetings list request
        meetings_response = MagicMock()
        meetings_response.status_code = 200
        meetings_response.json.return_value = {
            "value": [
                {
                    "id": "meeting-1",
                    "subject": "Team Meeting",
                    "startDateTime": "2026-01-10T09:00:00Z",
                    "joinWebUrl": "https://teams.microsoft.com/...",
                },
            ]
        }
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.side_effect = [user_response, meetings_response]
            
            result = await client.list_meetings()
            
            assert len(result) == 1
            assert result[0].meeting_id == "meeting-1"
            assert result[0].title == "Team Meeting"

    @pytest.mark.asyncio
    async def test_get_insights_not_found(self, mock_credential):
        """Should return placeholder when insights not available."""
        client = MeetingsClient(mock_credential)
        
        user_response = MagicMock()
        user_response.status_code = 200
        user_response.json.return_value = {"id": "user-123"}
        
        insights_response = MagicMock()
        insights_response.status_code = 404
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.side_effect = [user_response, insights_response]
            
            result = await client.get_insights("meeting-123")
            
            assert result.meeting_id == "meeting-123"
            assert len(result.notes) == 1
            assert "not yet available" in result.notes[0].text.lower()

    @pytest.mark.asyncio
    async def test_get_insights_success(self, mock_credential):
        """Should parse full insights response."""
        client = MeetingsClient(mock_credential)
        
        user_response = MagicMock()
        user_response.status_code = 200
        user_response.json.return_value = {"id": "user-123"}
        
        insights_response = MagicMock()
        insights_response.status_code = 200
        insights_response.json.return_value = {
            "value": [
                {
                    "subject": "Planning Meeting",
                    "meetingNotes": [
                        {"title": "Overview", "text": "Discussed roadmap"}
                    ],
                    "actionItems": [
                        {"title": "Draft spec", "text": "Write spec doc", "ownerDisplayName": "Alice"}
                    ],
                    "viewpoint": {
                        "mentionEvents": [
                            {
                                "eventDateTime": "2026-01-10T10:00:00Z",
                                "transcriptUtterance": "Can you review this?",
                                "speaker": {"displayName": "Bob"},
                            }
                        ]
                    },
                }
            ]
        }
        
        with patch.object(client, "_make_request", new_callable=AsyncMock) as mock_req:
            mock_req.side_effect = [user_response, insights_response]
            
            result = await client.get_insights("meeting-123")
            
            assert result.meeting_title == "Planning Meeting"
            assert len(result.notes) == 1
            assert len(result.action_items) == 1
            assert len(result.mentions) == 1

    def test_extract_meeting_id_from_url(self, mock_credential):
        """Should extract meeting ID from Teams URL."""
        client = MeetingsClient(mock_credential)
        
        url = "https://teams.microsoft.com/l/meetup-join/ABC123XYZ/0"
        result = client._extract_meeting_id(url)
        
        assert result == "ABC123XYZ"

    def test_extract_meeting_id_invalid_url(self, mock_credential):
        """Should return empty string for invalid URL."""
        client = MeetingsClient(mock_credential)
        
        result = client._extract_meeting_id("https://example.com")
        assert result == ""
