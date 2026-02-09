"""
Test WordNamespace functionality

测试 WordNamespace 的事件处理功能。

Architecture Note:
    WordNamespace 只处理 Client → Server 的事件通知（fire-and-forget）:
    - word:event:selectionChanged
    - word:event:documentModified

    Server → Client 的命令（word:get:*, word:insert:* 等）通过 MCP BaseTool
    + OfficeWorkspace.emit_to_document() 直接发送，不经过 namespace handler。
"""

import logging
from typing import Any

import pytest

from office4ai.environment.workspace.socketio.namespaces.word import WordNamespace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
)


class TestWordNamespace:
    """Test WordNamespace class"""

    @pytest.fixture
    def word_namespace(self) -> WordNamespace:
        """Create WordNamespace instance"""
        return WordNamespace()

    @pytest.fixture
    def connected_session(self, word_namespace: WordNamespace) -> Any:
        """Create a connected session for testing"""
        sid = "test_socket_123"
        connection_manager.register_client(sid, "client1", "file:///test.docx", "/word")
        yield sid
        # Cleanup
        connection_manager.unregister_client(sid)

    # ========================================================================
    # 初始化测试
    # ========================================================================

    def test_namespace_init(self, word_namespace: WordNamespace) -> None:
        """Test WordNamespace initializes with correct namespace"""
        assert word_namespace.namespace_name == "/word"

    def test_namespace_has_only_event_handlers(self, word_namespace: WordNamespace) -> None:
        """Test WordNamespace only has event handlers, not command handlers"""
        # 事件处理器应存在
        assert hasattr(word_namespace, "on_word_event_selectionChanged")
        assert hasattr(word_namespace, "on_word_event_documentModified")

        # 命令处理器不应存在（已移至 MCP BaseTool）
        assert not hasattr(word_namespace, "on_word_get_selectedContent")
        assert not hasattr(word_namespace, "on_word_get_visibleContent")
        assert not hasattr(word_namespace, "on_word_insert_text")
        assert not hasattr(word_namespace, "on_word_replace_selection")

    # ========================================================================
    # selectionChanged 事件测试
    # ========================================================================

    @pytest.mark.asyncio
    async def test_selection_changed_with_connected_client(
        self, word_namespace: WordNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test selectionChanged event with a registered client logs correctly"""
        data = {
            "eventType": "selectionChanged",
            "clientId": "client1",
            "documentUri": "file:///test.docx",
            "data": {"text": "Selected text", "length": 13},
            "timestamp": 1234567890,
        }

        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_event_selectionChanged(connected_session, data)

        assert any("Word selection changed" in record.message for record in caplog.records)
        assert any("text length: 13" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_selection_changed_with_unknown_client(
        self, word_namespace: WordNamespace, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test selectionChanged event with unregistered sid does not crash"""
        data = {
            "eventType": "selectionChanged",
            "clientId": "unknown",
            "documentUri": "file:///test.docx",
            "data": {"text": "", "length": 0},
            "timestamp": 1234567890,
        }

        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_event_selectionChanged("unknown_sid", data)

        # 未注册的 sid 不应产生日志（get_client_info 返回 None）
        assert not any("Word selection changed" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_selection_changed_with_empty_data(
        self, word_namespace: WordNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test selectionChanged event with missing data fields does not crash"""
        data: dict[str, Any] = {}

        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_event_selectionChanged(connected_session, data)

        assert any("text length: 0" in record.message for record in caplog.records)

    # ========================================================================
    # documentModified 事件测试
    # ========================================================================

    @pytest.mark.asyncio
    async def test_document_modified_with_connected_client(
        self, word_namespace: WordNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test documentModified event with a registered client logs correctly"""
        data = {
            "eventType": "documentModified",
            "clientId": "client1",
            "documentUri": "file:///test.docx",
            "data": {"modificationType": "insert"},
            "timestamp": 1234567890,
        }

        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_event_documentModified(connected_session, data)

        assert any("Word document modified" in record.message for record in caplog.records)
        assert any("type: insert" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_document_modified_with_unknown_client(
        self, word_namespace: WordNamespace, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test documentModified event with unregistered sid does not crash"""
        data = {
            "eventType": "documentModified",
            "clientId": "unknown",
            "documentUri": "file:///test.docx",
            "data": {"modificationType": "delete"},
            "timestamp": 1234567890,
        }

        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_event_documentModified("unknown_sid", data)

        assert not any("Word document modified" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_document_modified_with_empty_data(
        self, word_namespace: WordNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test documentModified event with missing data fields does not crash"""
        data: dict[str, Any] = {}

        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_event_documentModified(connected_session, data)

        assert any("type: None" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_document_modified_update_type(
        self, word_namespace: WordNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test documentModified event with update modification type"""
        data = {
            "eventType": "documentModified",
            "clientId": "client1",
            "documentUri": "file:///test.docx",
            "data": {"modificationType": "update"},
            "timestamp": 1234567890,
        }

        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_event_documentModified(connected_session, data)

        assert any("type: update" in record.message for record in caplog.records)
