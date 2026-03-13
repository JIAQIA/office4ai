"""
Contract Tests for ppt:get:slideLayouts

测试 ppt:get:slideLayouts 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_layouts_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功获取幻灯片版式列表。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_layouts_response(layout_count=5),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideLayouts", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideLayouts",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert len(result.data["layouts"]) == 5
        first_layout = result.data["layouts"][0]
        assert "id" in first_layout
        assert "name" in first_layout
        assert "type" in first_layout
        assert "placeholderCount" in first_layout

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "ppt:get:slideLayouts"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_layouts_with_options(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试带选项获取版式列表。"""

    def response_factory(request: dict) -> dict:
        assert "options" in request
        assert request["options"]["includePlaceholders"] is True

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_layouts_response(layout_count=3),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideLayouts", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideLayouts",
            params={
                "document_uri": client.document_uri,
                "options": {"includePlaceholders": True},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert len(result.data["layouts"]) == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_layouts_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Document not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/nonexistent.pptx",
    )

    client.register_response("ppt:get:slideLayouts", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideLayouts",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
