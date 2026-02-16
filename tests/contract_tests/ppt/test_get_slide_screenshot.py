"""
Contract Tests for ppt:get:slideScreenshot

测试 ppt:get:slideScreenshot 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_screenshot_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功获取幻灯片截图（PNG 格式）。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        assert request["slideIndex"] == 0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_screenshot_response(format="png"),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideScreenshot", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideScreenshot",
            params={"document_uri": client.document_uri, "slideIndex": 0},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["format"] == "png"
        assert len(result.data["base64"]) > 0

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "ppt:get:slideScreenshot"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_screenshot_jpeg_format(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试 JPEG 格式截图。"""

    def response_factory(request: dict) -> dict:
        assert "options" in request
        assert request["options"]["format"] == "jpeg"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_screenshot_response(format="jpeg"),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideScreenshot", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideScreenshot",
            params={
                "document_uri": client.document_uri,
                "slideIndex": 0,
                "options": {"format": "jpeg", "quality": 80},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["format"] == "jpeg"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_screenshot_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Slide not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideScreenshot", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideScreenshot",
            params={"document_uri": client.document_uri, "slideIndex": 999},
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
