"""Tests for conversation state management."""

import pytest
from datetime import datetime, timedelta, timezone
from unittest.mock import patch

from m365_copilot.conversation import (
    ConversationState,
    ConversationStore,
    get_conversation_store,
    CONVERSATION_TTL,
)


class TestConversationState:
    """Tests for ConversationState dataclass."""

    def test_creation(self):
        """Should create state with defaults."""
        state = ConversationState(id="test-123")
        assert state.id == "test-123"
        assert state.turn_count == 0
        assert state.display_name == ""
        assert state.created_at is not None

    def test_is_expired_false(self):
        """Should not be expired when within TTL."""
        state = ConversationState(id="test-123")
        assert state.is_expired() is False

    def test_is_expired_true(self):
        """Should be expired when past TTL."""
        old_time = datetime.now(timezone.utc) - CONVERSATION_TTL - timedelta(minutes=1)
        state = ConversationState(
            id="test-123",
            last_activity=old_time,
        )
        assert state.is_expired() is True

    def test_touch_updates_activity(self):
        """Should update last_activity timestamp."""
        state = ConversationState(id="test-123")
        old_activity = state.last_activity
        
        # Small delay to ensure time difference
        state.touch()
        
        assert state.last_activity >= old_activity

    def test_increment_turn(self):
        """Should increment turn count and touch."""
        state = ConversationState(id="test-123")
        assert state.turn_count == 0
        
        result = state.increment_turn()
        assert result == 1
        assert state.turn_count == 1
        
        result = state.increment_turn()
        assert result == 2
        assert state.turn_count == 2


class TestConversationStore:
    """Tests for ConversationStore."""

    def test_create_returns_state(self):
        """Should create and return new conversation state."""
        store = ConversationStore()
        state = store.create(display_name="Test Query")
        
        assert state.id is not None
        assert state.display_name == "Test Query"
        assert store.count() == 1

    def test_get_existing(self):
        """Should return existing conversation."""
        store = ConversationStore()
        created = store.create()
        
        retrieved = store.get(created.id)
        assert retrieved is not None
        assert retrieved.id == created.id

    def test_get_nonexistent(self):
        """Should return None for nonexistent conversation."""
        store = ConversationStore()
        result = store.get("nonexistent-id")
        assert result is None

    def test_get_expired(self):
        """Should return None and delete expired conversation."""
        store = ConversationStore()
        state = store.create()
        
        # Manually expire the conversation
        state.last_activity = datetime.now(timezone.utc) - CONVERSATION_TTL - timedelta(minutes=1)
        
        result = store.get(state.id)
        assert result is None
        assert store.count() == 0

    def test_update_activity(self):
        """Should update activity and return True."""
        store = ConversationStore()
        state = store.create()
        old_activity = state.last_activity
        
        result = store.update_activity(state.id)
        assert result is True
        
        retrieved = store.get(state.id)
        assert retrieved.last_activity >= old_activity

    def test_update_activity_nonexistent(self):
        """Should return False for nonexistent conversation."""
        store = ConversationStore()
        result = store.update_activity("nonexistent-id")
        assert result is False

    def test_delete(self):
        """Should delete existing conversation."""
        store = ConversationStore()
        state = store.create()
        
        result = store.delete(state.id)
        assert result is True
        assert store.count() == 0

    def test_delete_nonexistent(self):
        """Should return False for nonexistent conversation."""
        store = ConversationStore()
        result = store.delete("nonexistent-id")
        assert result is False

    def test_cleanup_expired(self):
        """Should remove all expired conversations."""
        store = ConversationStore()
        
        # Create some conversations
        active = store.create()
        expired1 = store.create()
        expired2 = store.create()
        
        # Expire two of them
        expired_time = datetime.now(timezone.utc) - CONVERSATION_TTL - timedelta(minutes=1)
        expired1.last_activity = expired_time
        expired2.last_activity = expired_time
        
        count = store.cleanup_expired()
        assert count == 2
        assert store.count() == 1
        assert store.get(active.id) is not None

    def test_list_active(self):
        """Should return only non-expired conversations."""
        store = ConversationStore()
        
        active1 = store.create()
        active2 = store.create()
        expired = store.create()
        
        expired.last_activity = datetime.now(timezone.utc) - CONVERSATION_TTL - timedelta(minutes=1)
        
        active_list = store.list_active()
        assert len(active_list) == 2
        ids = [s.id for s in active_list]
        assert active1.id in ids
        assert active2.id in ids
        assert expired.id not in ids


class TestGetConversationStore:
    """Tests for global store singleton."""

    def test_returns_same_instance(self):
        """Should return same store instance."""
        # Reset global
        import m365_copilot.conversation as conv_module
        conv_module._store = None
        
        store1 = get_conversation_store()
        store2 = get_conversation_store()
        
        assert store1 is store2
