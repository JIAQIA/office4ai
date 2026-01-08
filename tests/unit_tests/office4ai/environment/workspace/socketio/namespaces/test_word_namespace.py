"""
Test WordNamespace functionality

测试 WordNamespace 的所有核心功能。
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

    @pytest.mark.asyncio
    async def test_on_word_get_selected_content(
        self, word_namespace: WordNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test get selected content event handler (logging only)"""
        data = {
            "requestId": "req_123",
            "documentUri": "file:///test.docx",
            "options": {"includeText": True},
        }

        # Should log the request but not raise errors
        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_get_selectedContent(connected_session, data)

        # Verify logging occurred
        assert any("Received word:get:selectedContent from client1" in record.message for record in caplog.records)
        assert any("requestId: req_123" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_on_word_insert_text(
        self, word_namespace: WordNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test insert text event handler (logging only)"""
        data = {
            "requestId": "req_123",
            "documentUri": "file:///test.docx",
            "text": "Hello World",
            "location": "Cursor",
        }

        # Should log the request but not raise errors
        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_insert_text(connected_session, data)

        # Verify logging occurred
        assert any("Received word:insert:text from client1" in record.message for record in caplog.records)
        assert any("requestId: req_123" in record.message for record in caplog.records)
        assert any("text length: 11" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_on_word_replace_selection(
        self, word_namespace: WordNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test replace selection event handler (logging only)"""
        data = {
            "requestId": "req_123",
            "documentUri": "file:///test.docx",
            "content": {"text": "New content"},
        }

        # Should log the request but not raise errors
        with caplog.at_level(logging.INFO):
            await word_namespace.on_word_replace_selection(connected_session, data)

        # Verify logging occurred
        assert any("Received word:replace:selection from client1" in record.message for record in caplog.records)
        assert any("requestId: req_123" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_on_word_event_selection_changed(self, word_namespace: WordNamespace, connected_session: Any) -> None:
        """Test selection changed event"""
        data = {
            "eventType": "selectionChanged",
            "clientId": "client1",
            "documentUri": "file:///test.docx",
            "data": {"text": "Selected", "length": 8},
            "timestamp": 1234567890,
        }

        # Should log but not raise errors
        await word_namespace.on_word_event_selectionChanged(connected_session, data)

    @pytest.mark.asyncio
    async def test_on_word_event_document_modified(self, word_namespace: WordNamespace, connected_session: Any) -> None:
        """Test document modified event"""
        data = {
            "eventType": "documentModified",
            "clientId": "client1",
            "documentUri": "file:///test.docx",
            "data": {"modificationType": "insert"},
            "timestamp": 1234567890,
        }

        await word_namespace.on_word_event_documentModified(connected_session, data)

    @pytest.mark.asyncio
    async def test_unimplemented_events(self, word_namespace: WordNamespace, connected_session: Any) -> None:
        """Test unimplemented event handlers log warnings"""
        unimplemented_events = [
            "word:get:visibleContent",
            "word:get:documentStructure",
            "word:get:documentStats",
            "word:replace:text",
            "word:append:text",
            "word:insert:image",
            "word:insert:table",
            "word:insert:equation",
            "word:insert:toc",
            "word:export:content",
        ]

        for event in unimplemented_events:
            method_name = f"on_{event.replace(':', '_')}"
            method = getattr(word_namespace, method_name)
            data: dict[str, Any] = {"requestId": "req_123", "documentUri": "file:///test.docx"}

            # Should log warning but not crash
            await method(connected_session, data)
