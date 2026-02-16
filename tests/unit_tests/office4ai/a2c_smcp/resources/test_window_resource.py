"""WindowResource 单元测试 | WindowResource unit tests"""

from __future__ import annotations

import time
from unittest.mock import patch

import pytest

from office4ai.a2c_smcp.resources.window import WindowResource
from office4ai.environment.workspace.office_workspace import LastActivity, OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    ClientInfo,
    normalize_document_uri,
)


@pytest.fixture
def workspace() -> OfficeWorkspace:
    return OfficeWorkspace()


@pytest.fixture
def resource(workspace: OfficeWorkspace) -> WindowResource:
    return WindowResource(workspace, priority=0, fullscreen=True)


class TestWindowResourceMetadata:
    """Test URI format, base_uri, name, description, mime_type."""

    def test_uri_format(self, resource: WindowResource) -> None:
        assert resource.uri == "window://office4ai?priority=0&fullscreen=true"

    def test_base_uri(self, resource: WindowResource) -> None:
        assert resource.base_uri == "window://office4ai"

    def test_name_and_description(self, resource: WindowResource) -> None:
        assert resource.name
        assert resource.description
        assert isinstance(resource.name, str)
        assert isinstance(resource.description, str)

    def test_mime_type(self, resource: WindowResource) -> None:
        assert resource.mime_type == "text/plain"


class TestWindowResourceRead:
    """Test read() rendering with various workspace states."""

    @pytest.mark.asyncio
    async def test_read_no_documents(self, resource: WindowResource) -> None:
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = []
            content = await resource.read()

        assert "已连接文档 (0)" in content
        assert "暂无文档连接" in content

    @pytest.mark.asyncio
    async def test_read_with_documents(self, resource: WindowResource) -> None:
        clients = [
            ClientInfo(
                socket_id="s1",
                client_id="c1",
                document_uri="file:///tmp/report.docx",
                namespace="/word",
                connected_at=time.time(),
            ),
            ClientInfo(
                socket_id="s2",
                client_id="c2",
                document_uri="file:///tmp/draft.docx",
                namespace="/word",
                connected_at=time.time(),
            ),
        ]
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "已连接文档 (2)" in content
        assert "[word] file:///tmp/report.docx" in content
        assert "[word] file:///tmp/draft.docx" in content

    @pytest.mark.asyncio
    async def test_read_with_last_activity(self, resource: WindowResource) -> None:
        doc_uri = "file:///tmp/report.docx"
        normalized_uri = normalize_document_uri(doc_uri)
        resource.workspace._last_activity = LastActivity(
            document_uri=normalized_uri,
            tool_name="word_get_visible_content",
            timestamp=time.time() - 2,
        )
        resource.workspace._content_cache[normalized_uri] = "Hello world"
        resource.workspace._structure_cache[normalized_uri] = "Heading 1: Intro"

        clients = [
            ClientInfo(
                socket_id="s1",
                client_id="c1",
                document_uri=doc_uri,
                namespace="/word",
                connected_at=time.time(),
            ),
        ]
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert f"活跃文档: {normalized_uri}" in content
        assert "word_get_visible_content" in content
        assert "--- 可见内容 ---" in content
        assert "Hello world" in content
        assert "--- 文档结构 ---" in content
        assert "Heading 1: Intro" in content


class TestWindowResourceUpdateFromUri:
    """Test update_from_uri() parameter parsing."""

    def test_update_from_uri_priority(self, resource: WindowResource) -> None:
        resource.update_from_uri("window://office4ai?priority=50&fullscreen=true")
        assert resource.uri == "window://office4ai?priority=50&fullscreen=true"

    def test_update_from_uri_fullscreen(self, resource: WindowResource) -> None:
        resource.update_from_uri("window://office4ai?priority=0&fullscreen=false")
        assert resource.uri == "window://office4ai?priority=0&fullscreen=false"

    def test_update_from_uri_invalid_priority(self, resource: WindowResource) -> None:
        """Out-of-range priority is ignored."""
        resource.update_from_uri("window://office4ai?priority=200&fullscreen=true")
        # Priority should remain 0
        assert "priority=0" in resource.uri

    def test_priority_validation_in_constructor(self) -> None:
        ws = OfficeWorkspace()
        with pytest.raises(ValueError, match="priority must be int"):
            WindowResource(ws, priority=-1)
        with pytest.raises(ValueError, match="priority must be int"):
            WindowResource(ws, priority=101)
