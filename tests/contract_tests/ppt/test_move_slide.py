"""
Contract Tests for ppt:move:slide

测试 ppt:move:slide 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_move_slide_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功移动幻灯片。"""

    def response_factory(request: dict) -> dict:
        assert request["fromIndex"] == 2
        assert request["toIndex"] == 5

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.move_slide_response(from_index=2, to_index=5),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:move:slide", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="move:slide",
            params={
                "document_uri": client.document_uri,
                "fromIndex": 2,
                "toIndex": 5,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["fromIndex"] == 2
        assert result.data["toIndex"] == 5

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:move:slide"
        assert event_data["fromIndex"] == 2
        assert event_data["toIndex"] == 5
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_move_slide_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Invalid slide index"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:move:slide", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="move:slide",
            params={
                "document_uri": client.document_uri,
                "fromIndex": 999,
                "toIndex": 0,
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_move_slide_verify_indices(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试验证移动前后索引传递正确。"""

    def response_factory(request: dict) -> dict:
        # 验证从最后一个移到第一个
        assert request["fromIndex"] == 9
        assert request["toIndex"] == 0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.move_slide_response(from_index=9, to_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:move:slide", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="move:slide",
            params={
                "document_uri": client.document_uri,
                "fromIndex": 9,
                "toIndex": 0,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["fromIndex"] == 9
        assert result.data["toIndex"] == 0
    finally:
        await client.disconnect()
