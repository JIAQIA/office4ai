"""
Contract Tests for ppt:get:slideElements

测试 ppt:get:slideElements 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_elements_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功获取指定幻灯片元素。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        assert request["slideIndex"] == 3

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_elements_response(slide_index=3, element_count=5),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideElements", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideElements",
            params={"document_uri": client.document_uri, "slideIndex": 3},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["slideIndex"] == 3
        assert len(result.data["elements"]) == 5

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:get:slideElements"
        assert event_data["slideIndex"] == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_elements_with_options(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试带过滤选项获取元素。"""
    expected_options = {
        "includeText": True,
        "includeImages": False,
        "includeShapes": True,
        "includeTables": False,
        "includeCharts": False,
    }

    def response_factory(request: dict) -> dict:
        assert "options" in request
        actual_options = request["options"]
        for key, value in expected_options.items():
            assert actual_options.get(key) == value, f"Expected {key}={value}, got {actual_options.get(key)}"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_elements_response(slide_index=1, element_count=2),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideElements", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideElements",
            params={
                "document_uri": client.document_uri,
                "slideIndex": 1,
                "options": expected_options,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert len(result.data["elements"]) == 2
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_elements_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Slide index out of range"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideElements", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideElements",
            params={"document_uri": client.document_uri, "slideIndex": 999},
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3001" in str(result.error) or "out of range" in str(result.error).lower()
    finally:
        await client.disconnect()
