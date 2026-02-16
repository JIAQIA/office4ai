"""
Contract Tests for ppt:update:element

测试 ppt:update:element 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_element_position(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试更新元素位置。"""

    def response_factory(request: dict) -> dict:
        assert request["elementId"] == "shape-001"
        assert request["updates"]["left"] == 100.0
        assert request["updates"]["top"] == 200.0
        assert request["updates"]["width"] == 300.0
        assert request["updates"]["height"] == 150.0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_element_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:element", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:element",
            params={
                "document_uri": client.document_uri,
                "elementId": "shape-001",
                "updates": {"left": 100.0, "top": 200.0, "width": 300.0, "height": 150.0},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["updatedCount"] == 1

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:update:element"
        assert event_data["elementId"] == "shape-001"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_element_rotation_with_slide_index(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试更新元素旋转角度，并验证 slideIndex 传递。"""

    def response_factory(request: dict) -> dict:
        assert request["elementId"] == "shape-002"
        assert request["slideIndex"] == 3
        assert request["updates"]["rotation"] == 45.0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_element_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:element", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:element",
            params={
                "document_uri": client.document_uri,
                "elementId": "shape-002",
                "slideIndex": 3,
                "updates": {"rotation": 45.0},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_element_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Element not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:element", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:element",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
                "updates": {"left": 100.0},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
