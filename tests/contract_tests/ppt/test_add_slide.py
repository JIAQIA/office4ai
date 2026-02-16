"""
Contract Tests for ppt:add:slide

测试 ppt:add:slide 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_add_slide_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功添加幻灯片。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.add_slide_response(slide_index=5, slide_id="slide-006"),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:add:slide", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="add:slide",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["slideIndex"] == 5
        assert result.data["slideId"] == "slide-006"

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "ppt:add:slide"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_add_slide_with_layout_and_index(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试带版式和插入位置添加幻灯片。"""

    def response_factory(request: dict) -> dict:
        assert "options" in request
        assert request["options"]["insertIndex"] == 2
        assert request["options"]["layout"] == "Title Slide"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.add_slide_response(slide_index=2, slide_id="slide-new"),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:add:slide", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="add:slide",
            params={
                "document_uri": client.document_uri,
                "options": {"insertIndex": 2, "layout": "Title Slide"},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["slideIndex"] == 2
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_add_slide_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Failed to add slide"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:add:slide", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="add:slide",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
