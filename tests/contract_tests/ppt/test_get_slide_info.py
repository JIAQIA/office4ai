"""
Contract Tests for ppt:get:slideInfo

测试 ppt:get:slideInfo 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_info_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功获取演示文稿基本信息。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_info_response(slide_count=10, current_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideInfo", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideInfo",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["slideCount"] == 10
        assert result.data["currentSlideIndex"] == 0
        assert "dimensions" in result.data
        assert result.data["dimensions"]["width"] == 960.0
        assert result.data["dimensions"]["height"] == 540.0

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "ppt:get:slideInfo"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_info_with_slide_index(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试带 slideIndex 获取详细幻灯片信息。"""

    def response_factory(request: dict) -> dict:
        assert request.get("slideIndex") == 5

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_info_response(slide_count=10, current_index=5, with_slide_info=True),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideInfo", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideInfo",
            params={"document_uri": client.document_uri, "slideIndex": 5},
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert "slideInfo" in result.data
        assert result.data["slideInfo"]["layout"] == "Title Slide"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_info_empty_presentation(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试空演示文稿。"""
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/empty.pptx",
    )

    client.register_static_response(
        "ppt:get:slideInfo",
        {
            "requestId": "test_req_001",
            "success": True,
            "data": ppt_factory.slide_info_response(slide_count=0, current_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        },
    )

    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideInfo",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["slideCount"] == 0
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_info_error(
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

    client.register_response("ppt:get:slideInfo", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideInfo",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
