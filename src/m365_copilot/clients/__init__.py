"""Microsoft Graph API clients for M365 Copilot APIs.

All clients use the official Microsoft SDK (microsoft-agents-m365copilot-beta).
"""

from m365_copilot.clients.chat import ChatClient
from m365_copilot.clients.meetings import MeetingsClient
from m365_copilot.clients.retrieval import RetrievalClient
from m365_copilot.clients.search import SearchClient

__all__ = [
    "ChatClient",
    "MeetingsClient",
    "RetrievalClient",
    "SearchClient",
]
