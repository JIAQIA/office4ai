"""PptWindowResource 单元测试 | PptWindowResource unit tests"""

from __future__ import annotations

import asyncio
import time
from unittest.mock import AsyncMock, patch

import pytest

from office4ai.a2c_smcp.resources.ppt_window import PptWindowResource
from office4ai.environment.workspace.office_workspace import LastActivity, OfficeWorkspace
from tests.unit_tests.office4ai.a2c_smcp.resources.conftest import make_client


@pytest.fixture
def resource(workspace: OfficeWorkspace) -> PptWindowResource:
    return PptWindowResource(workspace, priority=50, fullscreen=True)


def _make_slide_info_emitter(
    slide_count: int = 20,
    current_index: int = 4,
    width: float = 960.0,
    height: float = 540.0,
    aspect_ratio: str = "16:9",
) -> AsyncMock:
    """Create a mock emit_to_document that handles slideInfo calls."""

    async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
        if "slideInfo" not in event:
            return {"success": False}

        slide_index = data.get("slide_index")
        if slide_index is None:
            return {
                "success": True,
                "data": {
                    "slideCount": slide_count,
                    "currentSlideIndex": current_index,
                    "dimensions": {"width": width, "height": height, "aspectRatio": aspect_ratio},
                },
            }
        else:
            return {
                "success": True,
                "data": {
                    "slideInfo": {
                        "title": f"幻灯片标题 {slide_index + 1}",
                        "notes": f"备注 {slide_index + 1}" if slide_index % 2 == 0 else "",
                    },
                    "elements": [
                        {"type": "TextBox", "id": f"tb-{slide_index}"},
                        {"type": "Image", "id": f"img-{slide_index}"},
                    ],
                },
            }

    return AsyncMock(side_effect=mock_emit)


class TestPptWindowResourceMetadata:
    def test_uri_format(self, resource: PptWindowResource) -> None:
        assert resource.uri == "window://office4ai/ppt?priority=50&fullscreen=true"

    def test_base_uri(self, resource: PptWindowResource) -> None:
        assert resource.base_uri == "window://office4ai/ppt"

    def test_name(self, resource: PptWindowResource) -> None:
        assert resource.name == "PPT 工作区"

    def test_update_from_uri_range(self, resource: PptWindowResource) -> None:
        resource.update_from_uri("window://office4ai/ppt?range=3&priority=50&fullscreen=true")
        assert resource._range == 3


class TestPptWindowResourceRead:
    @pytest.mark.asyncio
    async def test_read_with_active_presentation(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/presentation.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )
        resource.workspace.emit_to_document = _make_slide_info_emitter(slide_count=20, current_index=4)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "总张数: 20" in content
        assert "960" in content and "540" in content
        assert "当前幻灯片: 第 5 张" in content  # 0-based → 1-based
        assert "幻灯片摘要 (第 3-7 张)" in content
        assert "➡️" in content

    @pytest.mark.asyncio
    async def test_slide_range_at_beginning(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )
        resource.workspace.emit_to_document = _make_slide_info_emitter(slide_count=10, current_index=0)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "幻灯片摘要 (第 1-3 张)" in content

    @pytest.mark.asyncio
    async def test_slide_range_at_end(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )
        resource.workspace.emit_to_document = _make_slide_info_emitter(slide_count=20, current_index=19)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "幻灯片摘要 (第 18-20 张)" in content

    @pytest.mark.asyncio
    async def test_slide_range_small(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )
        resource.workspace.emit_to_document = _make_slide_info_emitter(slide_count=3, current_index=1)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "幻灯片摘要 (第 1-3 张)" in content

    @pytest.mark.asyncio
    async def test_slide_summary_fields(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )
        resource.workspace.emit_to_document = _make_slide_info_emitter(slide_count=5, current_index=2)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "TextBox×1" in content
        assert "Image×1" in content
        assert "备注 1" in content
        assert "备注: (无)" in content

    @pytest.mark.asyncio
    async def test_slide_no_notes(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )
        resource.workspace.emit_to_document = _make_slide_info_emitter(slide_count=3, current_index=1)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "备注: (无)" in content

    @pytest.mark.asyncio
    async def test_current_slide_marker(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )
        resource.workspace.emit_to_document = _make_slide_info_emitter(slide_count=5, current_index=2)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "➡️ 第 3 张" in content
        assert "(当前)" in content

    @pytest.mark.asyncio
    async def test_presentation_info_timeout(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )

        async def slow_emit(document_uri: str, event: str, data: dict) -> dict:
            await asyncio.sleep(10)
            return {"success": True, "data": {}}

        resource.workspace.emit_to_document = AsyncMock(side_effect=slow_emit)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            resource.FETCH_TIMEOUT = 0.1
            content = await resource.read()

        assert "元数据不可用" in content
        assert "幻灯片摘要" not in content

    @pytest.mark.asyncio
    async def test_partial_slide_timeout(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )

        async def mixed_emit(document_uri: str, event: str, data: dict) -> dict:
            slide_index = data.get("slide_index")
            if slide_index is None:
                return {
                    "success": True,
                    "data": {
                        "slideCount": 5,
                        "currentSlideIndex": 2,
                        "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
                    },
                }
            if slide_index == 1:
                await asyncio.sleep(10)
            return {
                "success": True,
                "data": {
                    "slideInfo": {"title": f"Slide {slide_index + 1}", "notes": ""},
                    "elements": [{"type": "TextBox", "id": "t1"}],
                },
            }

        resource.workspace.emit_to_document = AsyncMock(side_effect=mixed_emit)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            resource.FETCH_TIMEOUT = 0.1
            content = await resource.read()

        assert "幻灯片信息不可用" in content
        assert "Slide 1" in content

    @pytest.mark.asyncio
    async def test_all_slides_timeout(self, resource: PptWindowResource) -> None:
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )

        async def emit(document_uri: str, event: str, data: dict) -> dict:
            slide_index = data.get("slide_index")
            if slide_index is None:
                return {
                    "success": True,
                    "data": {
                        "slideCount": 3,
                        "currentSlideIndex": 1,
                        "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
                    },
                }
            await asyncio.sleep(10)
            return {"success": True, "data": {}}

        resource.workspace.emit_to_document = AsyncMock(side_effect=emit)

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            resource.FETCH_TIMEOUT = 0.1
            content = await resource.read()

        assert "总张数: 3" in content
        assert content.count("幻灯片信息不可用") == 3

    @pytest.mark.asyncio
    async def test_ignores_word_documents(self, resource: PptWindowResource) -> None:
        clients = [
            make_client("file:///tmp/pres.pptx", namespace="/ppt", socket_id="s1"),
            make_client("file:///tmp/report.docx", namespace="/word", socket_id="s2"),
        ]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "文档列表 (1)" in content
        assert "report.docx" not in content

    @pytest.mark.asyncio
    async def test_concurrent_slide_fetch(self, resource: PptWindowResource) -> None:
        """Verify ±N slides are fetched concurrently (not serially)."""
        doc_uri = "file:///tmp/pres.pptx"
        resource.workspace._last_activity = LastActivity(
            document_uri=doc_uri, tool_name="ppt_get_slide_info", timestamp=time.time()
        )

        async def slow_emit(document_uri: str, event: str, data: dict) -> dict:
            slide_index = data.get("slide_index")
            if slide_index is None:
                return {
                    "success": True,
                    "data": {
                        "slideCount": 10,
                        "currentSlideIndex": 5,
                        "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
                    },
                }
            await asyncio.sleep(0.2)  # Each slide takes 0.2s
            return {
                "success": True,
                "data": {
                    "slideInfo": {"title": f"Slide {slide_index + 1}", "notes": ""},
                    "elements": [],
                },
            }

        resource.workspace.emit_to_document = AsyncMock(side_effect=slow_emit)
        resource._range = 2  # 5 slides to fetch

        clients = [make_client(doc_uri, namespace="/ppt")]
        with patch("office4ai.a2c_smcp.resources.ppt_window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients

            start_time = asyncio.get_event_loop().time()
            content = await resource.read()
            elapsed = asyncio.get_event_loop().time() - start_time

        # If concurrent: ~0.2s. If serial (5 slides): ~1.0s.
        assert elapsed < 0.5, f"Expected concurrent fetch (<0.5s), got {elapsed:.2f}s"
        assert "Slide 4" in content
