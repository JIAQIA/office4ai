"""
Test BaseNamespace functionality

测试 BaseNamespace 的所有核心功能。
"""

from typing import Any
from unittest.mock import AsyncMock

import pytest

from office4ai.environment.workspace.socketio.namespaces.base import BaseNamespace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
)


class TestBaseNamespace:
    """Test BaseNamespace class"""

    @pytest.fixture
    def base_namespace(self) -> BaseNamespace:
        """Create BaseNamespace instance"""
        return BaseNamespace("/test")

    @pytest.mark.asyncio
    async def test_on_connect_success(
        self, base_namespace: BaseNamespace, valid_handshake_data: dict[str, Any]
    ) -> None:
        """Test successful connection"""
        sid = "test_socket_123"

        # Mock emit and disconnect
        base_namespace.emit = AsyncMock()  # type: ignore[method-assign]
        base_namespace.disconnect = AsyncMock()  # type: ignore[method-assign]

        await base_namespace.on_connect(sid, valid_handshake_data)

        # Verify client registered
        assert connection_manager.get_client_info(sid) is not None

        # Verify confirmation sent
        base_namespace.emit.assert_called_once()
        call_args = base_namespace.emit.call_args
        assert call_args[0][0] == "connection:established"
        assert call_args[1]["to"] == sid

        # Cleanup
        connection_manager.unregister_client(sid)

    @pytest.mark.asyncio
    async def test_on_connect_missing_client_id(
        self,
        base_namespace: BaseNamespace,
        invalid_handshake_data_missing_client_id: dict[str, Any],
    ) -> None:
        """Test connection fails without clientId"""
        sid = "test_socket_123"
        base_namespace.disconnect = AsyncMock()  # type: ignore[method-assign]

        await base_namespace.on_connect(sid, invalid_handshake_data_missing_client_id)

        # Should disconnect client
        base_namespace.disconnect.assert_called_once_with(sid)

    @pytest.mark.asyncio
    async def test_on_connect_missing_document_uri(
        self,
        base_namespace: BaseNamespace,
        invalid_handshake_data_missing_document_uri: dict[str, Any],
    ) -> None:
        """Test connection fails without documentUri"""
        sid = "test_socket_123"
        base_namespace.disconnect = AsyncMock()  # type: ignore[method-assign]

        await base_namespace.on_connect(sid, invalid_handshake_data_missing_document_uri)

        # Should disconnect client
        base_namespace.disconnect.assert_called_once_with(sid)

    @pytest.mark.asyncio
    async def test_on_disconnect(self, base_namespace: BaseNamespace, valid_handshake_data: dict[str, Any]) -> None:
        """Test client disconnection"""
        sid = "test_socket_123"

        # First connect
        base_namespace.emit = AsyncMock()  # type: ignore[method-assign]
        await base_namespace.on_connect(sid, valid_handshake_data)

        # Verify connected
        assert connection_manager.get_client_info(sid) is not None

        # Now disconnect
        await base_namespace.on_disconnect(sid)

        # Verify cleaned up
        assert connection_manager.get_client_info(sid) is None

    @pytest.mark.asyncio
    async def test_on_connection_status(self, base_namespace: BaseNamespace) -> None:
        """Test connection status updates (fire-and-forget)"""
        sid = "test_socket_123"
        status_data: dict[str, Any] = {"status": "ready", "timestamp": 1234567890}

        # Should not raise any errors
        await base_namespace.on_connection_status(sid, status_data)

    def test_get_client_info(self, base_namespace: BaseNamespace, valid_handshake_data: dict[str, Any]) -> None:
        """Test getting client info from namespace"""
        sid = "test_socket_123"

        # Manually register for test
        connection_manager.register_client(sid, "client1", "file:///test.docx", "/test")

        client = base_namespace.get_client_info(sid)
        assert client is not None
        assert client.socket_id == sid

        # Cleanup
        connection_manager.unregister_client(sid)
