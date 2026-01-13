"""Conversation state management for multi-turn M365 Copilot interactions.

Implements in-memory conversation storage with TTL-based cleanup (ADR-003).
"""

from __future__ import annotations

import logging
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from threading import Lock

logger = logging.getLogger(__name__)

# Conversation TTL - 1 hour (aligned with M365 API behavior)
CONVERSATION_TTL = timedelta(hours=1)


@dataclass
class ConversationState:
    """State for a single conversation with M365 Copilot."""

    id: str
    created_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    turn_count: int = 0
    last_activity: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    display_name: str = ""

    def is_expired(self) -> bool:
        """Check if conversation has exceeded TTL."""
        now = datetime.now(timezone.utc)
        return (now - self.last_activity) > CONVERSATION_TTL

    def touch(self) -> None:
        """Update last activity timestamp."""
        self.last_activity = datetime.now(timezone.utc)

    def increment_turn(self) -> int:
        """Increment turn count and return new value."""
        self.turn_count += 1
        self.touch()
        return self.turn_count


class ConversationStore:
    """Thread-safe in-memory conversation state storage.

    Conversations are stored in memory and cleaned up after TTL expiration.
    State is lost on server restart (acceptable for MCP stdio pattern per ADR-003).
    """

    def __init__(self) -> None:
        self._conversations: dict[str, ConversationState] = {}
        self._lock = Lock()

    def create(self, display_name: str = "") -> ConversationState:
        """Create a new conversation and return its state."""
        conversation_id = str(uuid.uuid4())
        state = ConversationState(id=conversation_id, display_name=display_name)

        with self._lock:
            self._conversations[conversation_id] = state
            logger.debug("Created conversation %s", conversation_id)

        return state

    def get(self, conversation_id: str) -> ConversationState | None:
        """Get conversation state by ID.

        Returns None if conversation doesn't exist or has expired.
        """
        with self._lock:
            state = self._conversations.get(conversation_id)
            if state is None:
                return None
            if state.is_expired():
                del self._conversations[conversation_id]
                logger.debug("Conversation %s expired", conversation_id)
                return None
            return state

    def update_activity(self, conversation_id: str) -> bool:
        """Update last activity timestamp for a conversation.

        Returns True if conversation exists and was updated.
        """
        with self._lock:
            state = self._conversations.get(conversation_id)
            if state is None or state.is_expired():
                return False
            state.touch()
            return True

    def delete(self, conversation_id: str) -> bool:
        """Delete a conversation.

        Returns True if conversation existed and was deleted.
        """
        with self._lock:
            if conversation_id in self._conversations:
                del self._conversations[conversation_id]
                logger.debug("Deleted conversation %s", conversation_id)
                return True
            return False

    def cleanup_expired(self) -> int:
        """Remove all expired conversations.

        Returns number of conversations cleaned up.
        """
        with self._lock:
            expired_ids = [
                cid for cid, state in self._conversations.items() if state.is_expired()
            ]
            for cid in expired_ids:
                del self._conversations[cid]
            if expired_ids:
                logger.info("Cleaned up %d expired conversations", len(expired_ids))
            return len(expired_ids)

    def count(self) -> int:
        """Return number of active conversations."""
        with self._lock:
            return len(self._conversations)

    def list_active(self) -> list[ConversationState]:
        """Return list of all active (non-expired) conversations."""
        with self._lock:
            return [
                state
                for state in self._conversations.values()
                if not state.is_expired()
            ]


# Global conversation store instance
_store: ConversationStore | None = None


def get_conversation_store() -> ConversationStore:
    """Get the global conversation store (singleton)."""
    global _store
    if _store is None:
        _store = ConversationStore()
    return _store
