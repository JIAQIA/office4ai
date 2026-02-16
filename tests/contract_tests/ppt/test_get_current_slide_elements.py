"""
Contract Tests for ppt:get:currentSlideElements

测试 ppt:get:currentSlideElements 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_current_slide_elements_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功获取当前幻灯片元素。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.current_slide_elements_response(slide_index=0, element_count=3),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:currentSlideElements", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:currentSlideElements",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["slideIndex"] == 0
        assert len(result.data["elements"]) == 3

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:get:currentSlideElements"
        assert "requestId" in event_data
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_current_slide_elements_empty_slide(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试空幻灯片（无元素）。"""
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/empty.pptx",
    )

    client.register_static_response(
        "ppt:get:currentSlideElements",
        {
            "requestId": "test_req_001",
            "success": True,
            "data": ppt_factory.current_slide_elements_response(slide_index=2, element_count=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        },
    )

    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:currentSlideElements",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["slideIndex"] == 2
        assert len(result.data["elements"]) == 0
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_current_slide_elements_error(
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

    client.register_response("ppt:get:currentSlideElements", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:currentSlideElements",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3001" in str(result.error) or "not found" in str(result.error).lower()
    finally:
        await client.disconnect()
