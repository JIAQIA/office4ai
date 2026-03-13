"""Activity tracking 单元测试 | Activity tracking unit tests"""

from __future__ import annotations

from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
    normalize_document_uri,
)


class TestActivityTracking:
    """Test OfficeWorkspace activity tracking and caching."""

    def setup_method(self) -> None:
        self.ws = OfficeWorkspace()
        self.doc_uri = "file:///tmp/test.docx"
        self.normalized_uri = normalize_document_uri(self.doc_uri)

    def test_update_last_activity(self) -> None:
        self.ws.update_last_activity(self.doc_uri, "word_insert_text", {"success": True})

        activity = self.ws.get_last_activity()
        assert activity is not None
        assert activity.document_uri == self.normalized_uri
        assert activity.tool_name == "word_insert_text"

    def test_content_cache_populated_by_get_visible_content(self) -> None:
        self.ws.update_last_activity(
            self.doc_uri,
            "word_get_visible_content",
            {"content": "Hello world"},
        )

        assert self.ws.get_cached_content(self.doc_uri) == "Hello world"
        # Structure cache should not be populated
        assert self.ws.get_cached_structure(self.doc_uri) is None

    def test_structure_cache_populated_by_get_document_structure(self) -> None:
        self.ws.update_last_activity(
            self.doc_uri,
            "word_get_document_structure",
            {"content": "Heading 1: Intro"},
        )

        assert self.ws.get_cached_structure(self.doc_uri) == "Heading 1: Intro"
        # Content cache should not be populated
        assert self.ws.get_cached_content(self.doc_uri) is None

    def test_cache_cleared_on_document_disconnect(self) -> None:
        """When a document fully disconnects, its caches should be cleared."""
        # Populate caches
        self.ws.update_last_activity(
            self.doc_uri,
            "word_get_visible_content",
            {"content": "cached text"},
        )
        assert self.ws.get_cached_content(self.doc_uri) == "cached text"

        # Register disconnect callback (simulates what start() does)
        connection_manager.register_disconnect_callback(self.ws._clear_document_cache)

        # Simulate a client connecting and then disconnecting
        connection_manager.register_client(
            socket_id="test-sid",
            client_id="test-cid",
            document_uri=self.doc_uri,
            namespace="/word",
        )
        connection_manager.unregister_client("test-sid")

        # Caches should be cleared
        assert self.ws.get_cached_content(self.doc_uri) is None
        assert self.ws.get_last_activity() is None

    def teardown_method(self) -> None:
        # Clean up global connection_manager state
        connection_manager._on_document_disconnect.clear()
        for sid in list(connection_manager._clients.keys()):
            connection_manager.unregister_client(sid)
