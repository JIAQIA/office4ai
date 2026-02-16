"""
Contract Tests for ppt:reorder:element

测试 ppt:reorder:element 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_reorder_element_bring_to_front(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试 bringToFront 操作。"""

    def response_factory(request: dict) -> dict:
        assert request["elementId"] == "shape-001"
        assert request["action"] == "bringToFront"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.reorder_element_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:reorder:element", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="reorder:element",
            params={
                "document_uri": client.document_uri,
                "elementId": "shape-001",
                "action": "bringToFront",
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["reordered"] is True

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:reorder:element"
        assert event_data["action"] == "bringToFront"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_reorder_element_send_to_back(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试 sendToBack 操作。"""

    def response_factory(request: dict) -> dict:
        assert request["action"] == "sendToBack"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.reorder_element_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:reorder:element", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="reorder:element",
            params={
                "document_uri": client.document_uri,
                "elementId": "shape-002",
                "action": "sendToBack",
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_reorder_element_error(
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

    client.register_response("ppt:reorder:element", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="reorder:element",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
                "action": "bringToFront",
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
