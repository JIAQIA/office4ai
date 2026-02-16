"""
Contract Tests for ppt:insert:shape

测试 ppt:insert:shape 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_shape_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功插入形状。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        assert request["shapeType"] == "Rectangle"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_shape_response(shape_id="shape-020", slide_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:shape", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:shape",
            params={"document_uri": client.document_uri, "shapeType": "Rectangle"},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["elementId"] == "shape-020"

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:insert:shape"
        assert event_data["shapeType"] == "Rectangle"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_shape_different_types(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试不同形状类型。"""

    def response_factory(request: dict) -> dict:
        assert request["shapeType"] == "Circle"
        assert "options" in request
        assert request["options"]["fillColor"] == "#0000FF"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_shape_response(shape_id="shape-021", slide_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:shape", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:shape",
            params={
                "document_uri": client.document_uri,
                "shapeType": "Circle",
                "options": {
                    "left": 200.0,
                    "top": 150.0,
                    "width": 100.0,
                    "height": 100.0,
                    "fillColor": "#0000FF",
                },
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["elementId"] == "shape-021"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_shape_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Shape insert failed"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:shape", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:shape",
            params={"document_uri": client.document_uri, "shapeType": "Triangle"},
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
