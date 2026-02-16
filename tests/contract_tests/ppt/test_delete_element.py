"""
Contract Tests for ppt:delete:element

测试 ppt:delete:element 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_delete_element_single(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试删除单个元素。"""

    def response_factory(request: dict) -> dict:
        assert request["elementId"] == "shape-001"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.delete_element_response(deleted_count=1),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:delete:element", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="delete:element",
            params={
                "document_uri": client.document_uri,
                "elementId": "shape-001",
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["deletedCount"] == 1

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:delete:element"
        assert event_data["elementId"] == "shape-001"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_delete_element_batch(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试批量删除元素。"""
    ids_to_delete = ["shape-001", "shape-002", "shape-003"]

    def response_factory(request: dict) -> dict:
        assert request["elementIds"] == ids_to_delete

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.delete_element_response(deleted_count=3),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:delete:element", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="delete:element",
            params={
                "document_uri": client.document_uri,
                "elementIds": ids_to_delete,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["deletedCount"] == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_delete_element_error(
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

    client.register_response("ppt:delete:element", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="delete:element",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
