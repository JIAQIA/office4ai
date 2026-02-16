"""
Contract Tests for ppt:insert:table

测试 ppt:insert:table 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_table_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功插入表格。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        assert "options" in request
        assert request["options"]["rows"] == 3
        assert request["options"]["columns"] == 4

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_table_response(element_id="shape-030", rows=3, columns=4),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:table", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:table",
            params={
                "document_uri": client.document_uri,
                "options": {"rows": 3, "columns": 4},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["elementId"] == "shape-030"
        assert result.data["rows"] == 3
        assert result.data["columns"] == 4

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:insert:table"
        assert event_data["options"]["rows"] == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_table_with_data(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试带初始数据插入表格。"""
    table_data = [["A1", "B1", "C1"], ["A2", "B2", "C2"]]

    def response_factory(request: dict) -> dict:
        assert request["options"]["data"] == table_data
        assert request["options"]["rows"] == 2
        assert request["options"]["columns"] == 3

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_table_response(element_id="shape-031", rows=2, columns=3),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:table", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:table",
            params={
                "document_uri": client.document_uri,
                "options": {"rows": 2, "columns": 3, "data": table_data},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["rows"] == 2
        assert result.data["columns"] == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_table_with_position(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试带位置选项插入表格。"""

    def response_factory(request: dict) -> dict:
        opts = request["options"]
        assert opts["slideIndex"] == 2
        assert opts["left"] == 50.0
        assert opts["top"] == 100.0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_table_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:table", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:table",
            params={
                "document_uri": client.document_uri,
                "options": {"rows": 3, "columns": 4, "slideIndex": 2, "left": 50.0, "top": 100.0},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_table_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Table insert failed"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:table", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:table",
            params={
                "document_uri": client.document_uri,
                "options": {"rows": 3, "columns": 4},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
