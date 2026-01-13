"""Microsoft Graph API clients for M365 Copilot APIs."""

from m365_copilot.clients.base import GraphClient
from m365_copilot.clients.chat import ChatClient
from m365_copilot.clients.meetings import MeetingsClient
from m365_copilot.clients.retrieval import RetrievalClient
from m365_copilot.clients.search import SearchClient

__all__ = [
    "GraphClient",
    "ChatClient",
    "MeetingsClient",
    "RetrievalClient",
    "SearchClient",
]
