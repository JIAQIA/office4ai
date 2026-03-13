"""MCP Resource Subscription Manager.

Tracks which ServerSessions have subscribed to which resource URIs,
and sends ``resource_updated`` notifications when content changes.
"""

from __future__ import annotations

import asyncio
import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from mcp.server.session import ServerSession

logger = logging.getLogger(__name__)


class SubscriptionManager:
    """Manages resource subscriptions and dispatches update notifications."""

    def __init__(self) -> None:
        self._subscriptions: dict[str, set[ServerSession]] = {}

    def subscribe(self, uri: str, session: ServerSession) -> None:
        """Register *session* as a subscriber of *uri*."""
        self._subscriptions.setdefault(uri, set()).add(session)
        logger.debug("Session subscribed to %s (total: %d)", uri, len(self._subscriptions[uri]))

    def unsubscribe(self, uri: str, session: ServerSession) -> None:
        """Remove *session* from subscribers of *uri*."""
        sessions = self._subscriptions.get(uri)
        if sessions:
            sessions.discard(session)
            if not sessions:
                del self._subscriptions[uri]
        logger.debug("Session unsubscribed from %s", uri)

    async def notify(self, uri: str) -> None:
        """Send ``resource_updated`` to all sessions subscribed to *uri*.

        Dead sessions (broken pipe, closed connection) are automatically
        removed on failure (lazy cleanup).
        """
        from pydantic import AnyUrl

        sessions = self._subscriptions.get(uri)
        if not sessions:
            return

        dead: list[ServerSession] = []
        any_url = AnyUrl(uri)

        for session in sessions:
            try:
                await session.send_resource_updated(any_url)
            except Exception:
                logger.debug("Failed to notify session for %s, removing dead session", uri)
                dead.append(session)

        for s in dead:
            sessions.discard(s)
        if not sessions:
            del self._subscriptions[uri]

    async def notify_many(self, uris: list[str]) -> None:
        """Send ``resource_updated`` for each URI in *uris*."""
        for uri in uris:
            await self.notify(uri)

    def notify_fire_and_forget(self, uris: list[str]) -> None:
        """Schedule :meth:`notify_many` from a synchronous context.

        Uses the running event loop's ``create_task`` to bridge sync → async.
        """
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            logger.warning("No running event loop; skipping subscription notification")
            return
        loop.create_task(self.notify_many(uris))

    def clear(self) -> None:
        """Remove all subscriptions."""
        self._subscriptions.clear()
        logger.debug("All subscriptions cleared")
