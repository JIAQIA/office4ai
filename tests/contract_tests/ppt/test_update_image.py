"""
Contract Tests for ppt:update:image

测试 ppt:update:image 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

TEST_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_image_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功替换图片内容。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert request["elementId"] == "img-001"
        assert request["image"]["base64"] == TEST_BASE64

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_image_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:image", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:image",
            params={
                "document_uri": client.document_uri,
                "elementId": "img-001",
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["updatedCount"] == 1

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:update:image"
        assert event_data["elementId"] == "img-001"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_image_keep_dimensions(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试 keepDimensions 选项传递。"""

    def response_factory(request: dict) -> dict:
        assert "options" in request
        assert request["options"]["keepDimensions"] is False
        assert request["options"]["width"] == 500.0
        assert request["options"]["height"] == 300.0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_image_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:image", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:image",
            params={
                "document_uri": client.document_uri,
                "elementId": "img-001",
                "image": {"base64": TEST_BASE64},
                "options": {"keepDimensions": False, "width": 500.0, "height": 300.0},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_image_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Image element not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:image", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:image",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
