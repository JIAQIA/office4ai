"""
Contract Tests for ppt:insert:image

测试 ppt:insert:image 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

TEST_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_image_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功插入图片。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        assert "image" in request
        assert request["image"]["base64"] == TEST_BASE64

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_image_response(image_id="shape-025", slide_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:image", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:image",
            params={
                "document_uri": client.document_uri,
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["elementId"] == "shape-025"

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:insert:image"
        assert event_data["image"]["base64"] == TEST_BASE64
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_image_with_position_options(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试带位置选项插入图片。"""

    def response_factory(request: dict) -> dict:
        assert "options" in request
        opts = request["options"]
        assert opts["left"] == 100.0
        assert opts["top"] == 200.0
        assert opts["slideIndex"] == 1

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_image_response(image_id="shape-026", slide_index=1),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:image", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:image",
            params={
                "document_uri": client.document_uri,
                "image": {"base64": TEST_BASE64},
                "options": {"slideIndex": 1, "left": 100.0, "top": 200.0, "width": 400.0, "height": 300.0},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["slideIndex"] == 1
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_image_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Image insert failed"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:image", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:image",
            params={
                "document_uri": client.document_uri,
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
