"""WordWindowResource 单元测试 | WordWindowResource unit tests"""

from __future__ import annotations

import asyncio
import time
from unittest.mock import AsyncMock, patch

import pytest

from office4ai.a2c_smcp.resources.word_window import WordWindowResource
from office4ai.environment.workspace.office_workspace import LastActivity, OfficeWorkspace
from tests.unit_tests.office4ai.a2c_smcp.resources.conftest import make_client


@pytest.fixture
def resource(workspace: OfficeWorkspace) -> WordWindowResource:
    return WordWindowResource(workspace, priority=50, fullscreen=True)


class TestWordWindowResourceMetadata:
    def test_uri_format(self, resource: WordWindowResource) -> None:
        assert resource.uri == "window://office4ai/word?priority=50&fullscreen=true"

    def test_base_uri(self, resource: WordWindowResource) -> None:
        assert resource.base_uri == "window://office4ai/word"

    def test_name(self, resource: WordWindowResource) -> None:
        assert resource.name == "Word 工作区"

    def test_update_from_uri_priority(self, resource: WordWindowResource) -> None:
        resource.update_from_uri("window://office4ai/word?priority=80&fullscreen=true")
        assert "priority=80" in resource.uri


class TestWordWindowResourceRead:
    @pytest.mark.asyncio
    async def test_read_no_documents(self, resource: WordWindowResource) -> None:
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = []
            content = await resource.read()

        assert "文档列表 (0)" in content
        assert "暂无 Word 文档连接" in content

    @pytest.mark.asyncio
    async def test_read_documents_no_active(self, resource: WordWindowResource) -> None:
        clients = [make_client("file:///tmp/report.docx")]
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "文档列表 (1)" in content
        assert "file:///tmp/report.docx" in content
        assert "激活文档" not in content

    @pytest.mark.asyncio
    async def test_read_with_active_document(self, resource: WordWindowResource) -> None:
        doc_uri = "file:///tmp/report.docx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="word_get_visible_content", timestamp=time.time()
        )

        clients = [make_client(doc_uri)]

        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "documentStats" in event:
                return {"success": True, "data": {"pageCount": 5, "wordCount": 1200, "paragraphCount": 20}}
            if "visibleContent" in event:
                return {"success": True, "data": {"text": "Hello World\nParagraph 2"}}
            return {"success": False}

        resource.workspace.emit_to_document = AsyncMock(side_effect=mock_emit)

        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "⭐" in content
        assert "(激活)" in content
        assert "激活文档: report.docx" in content
        assert "总页数: 5" in content
        assert "总字数: 1,200" in content
        assert "段落数: 20" in content
        assert "Hello World" in content

    @pytest.mark.asyncio
    async def test_read_multiple_documents(self, resource: WordWindowResource) -> None:
        doc1 = "file:///tmp/report.docx"
        doc2 = "file:///tmp/draft.docx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc1, tool_name="word_get_visible_content", timestamp=time.time()
        )

        clients = [
            make_client(doc1, socket_id="s1", client_id="c1"),
            make_client(doc2, socket_id="s2", client_id="c2"),
        ]

        resource.workspace.emit_to_document = AsyncMock(
            return_value={"success": True, "data": {"pageCount": 1, "wordCount": 10, "paragraphCount": 1}}
        )

        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "文档列表 (2)" in content
        assert "⭐" in content
        assert doc2 in content

    @pytest.mark.asyncio
    async def test_ignores_ppt_documents(self, resource: WordWindowResource) -> None:
        clients = [
            make_client("file:///tmp/report.docx", namespace="/word", socket_id="s1"),
            make_client("file:///tmp/slides.pptx", namespace="/ppt", socket_id="s2"),
        ]
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "文档列表 (1)" in content
        assert "slides.pptx" not in content

    @pytest.mark.asyncio
    async def test_ignores_ppt_active(self, resource: WordWindowResource) -> None:
        """last_activity 指向 /ppt 文档 → Word 视图无激活文档"""
        resource.workspace._last_activity = LastActivity(
            document_uri="file:///tmp/slides.pptx", tool_name="ppt_get_slide_info", timestamp=time.time()
        )
        clients = [
            make_client("file:///tmp/report.docx", namespace="/word"),
            make_client("file:///tmp/slides.pptx", namespace="/ppt", socket_id="s2", client_id="c2"),
        ]
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "文档列表 (1)" in content
        assert "激活文档" not in content

    @pytest.mark.asyncio
    async def test_stats_timeout(self, resource: WordWindowResource) -> None:
        doc_uri = "file:///tmp/report.docx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="word_get_visible_content", timestamp=time.time()
        )

        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "documentStats" in event:
                await asyncio.sleep(10)  # will be cancelled by timeout
            return {"success": True, "data": {"text": "Hello"}}

        resource.workspace.emit_to_document = AsyncMock(side_effect=mock_emit)

        clients = [make_client(doc_uri)]
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            resource.FETCH_TIMEOUT = 0.1
            content = await resource.read()

        assert "元数据不可用" in content
        assert "Hello" in content

    @pytest.mark.asyncio
    async def test_visible_content_timeout(self, resource: WordWindowResource) -> None:
        doc_uri = "file:///tmp/report.docx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="word_get_visible_content", timestamp=time.time()
        )

        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "visibleContent" in event:
                await asyncio.sleep(10)
            return {"success": True, "data": {"pageCount": 5, "wordCount": 100, "paragraphCount": 3}}

        resource.workspace.emit_to_document = AsyncMock(side_effect=mock_emit)

        clients = [make_client(doc_uri)]
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            resource.FETCH_TIMEOUT = 0.1
            content = await resource.read()

        assert "总页数: 5" in content
        assert "可见内容不可用" in content

    @pytest.mark.asyncio
    async def test_all_timeout(self, resource: WordWindowResource) -> None:
        doc_uri = "file:///tmp/report.docx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="word_get_visible_content", timestamp=time.time()
        )

        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            await asyncio.sleep(10)
            return {"success": True, "data": {}}

        resource.workspace.emit_to_document = AsyncMock(side_effect=mock_emit)

        clients = [make_client(doc_uri)]
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            resource.FETCH_TIMEOUT = 0.1
            content = await resource.read()

        assert "文档列表 (1)" in content
        assert "元数据不可用" in content
        assert "可见内容不可用" in content

    @pytest.mark.asyncio
    async def test_concurrent_fetch(self, resource: WordWindowResource) -> None:
        """Verify stats and visibleContent are fetched concurrently (not serially)."""
        doc_uri = "file:///tmp/report.docx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="word_get_visible_content", timestamp=time.time()
        )

        async def slow_emit(document_uri: str, event: str, data: dict) -> dict:
            await asyncio.sleep(0.2)  # Each takes 0.2s
            if "documentStats" in event:
                return {"success": True, "data": {"pageCount": 1, "wordCount": 10, "paragraphCount": 1}}
            return {"success": True, "data": {"text": "content"}}

        resource.workspace.emit_to_document = AsyncMock(side_effect=slow_emit)

        clients = [make_client(doc_uri)]
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients

            start_time = asyncio.get_event_loop().time()
            content = await resource.read()
            elapsed = asyncio.get_event_loop().time() - start_time

        # If concurrent: ~0.2s. If serial: ~0.4s.
        assert elapsed < 0.35, f"Expected concurrent fetch (<0.35s), got {elapsed:.2f}s"
        assert "总页数: 1" in content
        assert "content" in content

    @pytest.mark.asyncio
    async def test_read_empty_visible_content(self, resource: WordWindowResource) -> None:
        """visibleContent returns empty text → renders '(空)'."""
        doc_uri = "file:///tmp/empty.docx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="word_get_visible_content", timestamp=time.time()
        )

        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "documentStats" in event:
                return {"success": True, "data": {"pageCount": 1, "wordCount": 0, "paragraphCount": 0}}
            return {"success": True, "data": {"text": ""}}

        resource.workspace.emit_to_document = AsyncMock(side_effect=mock_emit)

        clients = [make_client(doc_uri)]
        with patch("office4ai.a2c_smcp.resources.word_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "(空)" in content
