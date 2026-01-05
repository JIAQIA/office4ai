"""
Test ConnectionManager functionality

测试 ConnectionManager 的所有核心功能。
"""

import time

import pytest

from office4ai.environment.workspace.socketio.services.connection_manager import (
    ConnectionManager,
    ClientInfo,
)


class TestConnectionManager:
    """Test ConnectionManager class"""

    def test_initialization(self, connection_manager: ConnectionManager) -> None:
        """Test manager initializes with empty state"""
        assert connection_manager.get_connection_count() == 0
        assert connection_manager.get_document_count() == 0
        assert len(connection_manager._clients) == 0
        assert len(connection_manager._document_to_sockets) == 0

    def test_register_client(
        self, connection_manager: ConnectionManager, mock_client_info: ClientInfo
    ) -> None:
        """Test registering a new client"""
        client = connection_manager.register_client(
            socket_id=mock_client_info.socket_id,
            client_id=mock_client_info.client_id,
            document_uri=mock_client_info.document_uri,
            namespace=mock_client_info.namespace,
        )

        assert client.socket_id == mock_client_info.socket_id
        assert client.client_id == mock_client_info.client_id
        assert connection_manager.get_connection_count() == 1
        assert connection_manager.get_document_count() == 1

    def test_register_multiple_clients_same_document(
        self, connection_manager: ConnectionManager
    ) -> None:
        """Test multiple clients working on same document"""
        connection_manager.register_client("socket1", "client1", "file:///test.docx", "/word")
        connection_manager.register_client("socket2", "client2", "file:///test.docx", "/word")

        assert connection_manager.get_connection_count() == 2
        assert connection_manager.get_document_count() == 1
        assert len(connection_manager.get_clients_by_document("file:///test.docx")) == 2

    def test_unregister_client(
        self, connection_manager: ConnectionManager, mock_client_info: ClientInfo
    ) -> None:
        """Test unregistering a client"""
        connection_manager.register_client(
            mock_client_info.socket_id,
            mock_client_info.client_id,
            mock_client_info.document_uri,
            mock_client_info.namespace,
        )

        unregistered = connection_manager.unregister_client(mock_client_info.socket_id)

        assert unregistered is not None
        assert unregistered.client_id == mock_client_info.client_id
        assert connection_manager.get_connection_count() == 0
        assert connection_manager.get_document_count() == 0

    def test_unregister_nonexistent_client(self, connection_manager: ConnectionManager) -> None:
        """Test unregistering a client that doesn't exist"""
        result = connection_manager.unregister_client("nonexistent_socket")
        assert result is None

    def test_get_socket_by_document(self, connection_manager: ConnectionManager) -> None:
        """Test getting socket ID for a document"""
        connection_manager.register_client("socket1", "client1", "file:///test.docx", "/word")

        socket_id = connection_manager.get_socket_by_document("file:///test.docx")
        assert socket_id == "socket1"

        socket_id = connection_manager.get_socket_by_document("file:///nonexistent.docx")
        assert socket_id is None

    def test_get_client_info(self, connection_manager: ConnectionManager) -> None:
        """Test getting client information"""
        connection_manager.register_client("socket1", "client1", "file:///test.docx", "/word")

        client = connection_manager.get_client_info("socket1")
        assert client is not None
        assert client.client_id == "client1"

        client = connection_manager.get_client_info("nonexistent")
        assert client is None

    def test_get_clients_by_namespace(self, connection_manager: ConnectionManager) -> None:
        """Test getting clients by namespace"""
        connection_manager.register_client("socket1", "client1", "file:///test.docx", "/word")
        connection_manager.register_client("socket2", "client2", "file:///test.pptx", "/ppt")

        word_clients = connection_manager.get_clients_by_namespace("/word")
        ppt_clients = connection_manager.get_clients_by_namespace("/ppt")

        assert len(word_clients) == 1
        assert len(ppt_clients) == 1
        assert word_clients[0].socket_id == "socket1"
        assert ppt_clients[0].socket_id == "socket2"

    def test_is_document_active(self, connection_manager: ConnectionManager) -> None:
        """Test checking if document is active"""
        connection_manager.register_client("socket1", "client1", "file:///test.docx", "/word")

        assert connection_manager.is_document_active("file:///test.docx") is True
        assert connection_manager.is_document_active("file:///other.docx") is False

    def test_multiple_documents_cleanup(self, connection_manager: ConnectionManager) -> None:
        """Test document cleanup when all clients disconnect"""
        connection_manager.register_client("socket1", "client1", "file:///test.docx", "/word")
        connection_manager.register_client("socket2", "client2", "file:///test.docx", "/word")

        assert connection_manager.is_document_active("file:///test.docx")

        connection_manager.unregister_client("socket1")
        assert connection_manager.is_document_active("file:///test.docx")

        connection_manager.unregister_client("socket2")
        assert not connection_manager.is_document_active("file:///test.docx")
