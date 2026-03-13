"""Tests for SubscriptionManager."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.subscriptions import SubscriptionManager


def _make_session() -> MagicMock:
    """Create a mock ServerSession with send_resource_updated."""
    session = MagicMock()
    session.send_resource_updated = AsyncMock()
    return session


class TestSubscriptionManager:
    def setup_method(self) -> None:
        self.mgr = SubscriptionManager()

    def test_subscribe_and_unsubscribe(self) -> None:
        session = _make_session()
        self.mgr.subscribe("window://office4ai/word", session)
        assert "window://office4ai/word" in self.mgr._subscriptions
        assert session in self.mgr._subscriptions["window://office4ai/word"]

        self.mgr.unsubscribe("window://office4ai/word", session)
        assert "window://office4ai/word" not in self.mgr._subscriptions

    def test_unsubscribe_nonexistent(self) -> None:
        """Unsubscribing a URI that was never subscribed should not raise."""
        session = _make_session()
        self.mgr.unsubscribe("window://office4ai/word", session)

    @pytest.mark.asyncio
    async def test_notify_sends_to_all_subscribers(self) -> None:
        s1 = _make_session()
        s2 = _make_session()
        uri = "window://office4ai/word"

        self.mgr.subscribe(uri, s1)
        self.mgr.subscribe(uri, s2)

        await self.mgr.notify(uri)

        s1.send_resource_updated.assert_awaited_once()
        s2.send_resource_updated.assert_awaited_once()

    @pytest.mark.asyncio
    async def test_notify_no_subscribers(self) -> None:
        """notify on an unsubscribed URI should be a no-op."""
        await self.mgr.notify("window://office4ai/word")

    @pytest.mark.asyncio
    async def test_notify_removes_dead_session(self) -> None:
        alive = _make_session()
        dead = _make_session()
        dead.send_resource_updated.side_effect = Exception("connection lost")
        uri = "window://office4ai/word"

        self.mgr.subscribe(uri, alive)
        self.mgr.subscribe(uri, dead)

        await self.mgr.notify(uri)

        # Dead session should have been removed
        assert dead not in self.mgr._subscriptions.get(uri, set())
        # Alive session remains
        assert alive in self.mgr._subscriptions[uri]

    @pytest.mark.asyncio
    async def test_notify_removes_uri_when_all_sessions_dead(self) -> None:
        dead = _make_session()
        dead.send_resource_updated.side_effect = Exception("closed")
        uri = "window://office4ai/ppt"

        self.mgr.subscribe(uri, dead)
        await self.mgr.notify(uri)

        assert uri not in self.mgr._subscriptions

    @pytest.mark.asyncio
    async def test_notify_many(self) -> None:
        s = _make_session()
        self.mgr.subscribe("window://office4ai/word", s)
        self.mgr.subscribe("window://office4ai", s)

        await self.mgr.notify_many(["window://office4ai/word", "window://office4ai"])

        assert s.send_resource_updated.await_count == 2

    def test_clear(self) -> None:
        s = _make_session()
        self.mgr.subscribe("window://office4ai/word", s)
        self.mgr.subscribe("window://office4ai/ppt", s)

        self.mgr.clear()
        assert len(self.mgr._subscriptions) == 0

    def test_multiple_sessions_same_uri(self) -> None:
        s1 = _make_session()
        s2 = _make_session()
        uri = "window://office4ai/word"

        self.mgr.subscribe(uri, s1)
        self.mgr.subscribe(uri, s2)

        # Unsubscribe one; the other should remain
        self.mgr.unsubscribe(uri, s1)
        assert uri in self.mgr._subscriptions
        assert s2 in self.mgr._subscriptions[uri]
