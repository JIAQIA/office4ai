"""
Contract Tests for ppt:update:tableRowColumn

测试 ppt:update:tableRowColumn 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_row_column_rows(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试按行批量更新。"""
    rows_data = [
        {"rowIndex": 0, "values": ["A1", "B1", "C1"]},
        {"rowIndex": 1, "values": ["A2", "B2", "C2"]},
    ]

    def response_factory(request: dict) -> dict:
        assert request["elementId"] == "table-001"
        assert len(request["rows"]) == 2
        assert request["rows"][0]["values"] == ["A1", "B1", "C1"]

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_table_row_column_response(updated_count=2),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableRowColumn", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableRowColumn",
            params={
                "document_uri": client.document_uri,
                "elementId": "table-001",
                "rows": rows_data,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["updatedCount"] == 2

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "ppt:update:tableRowColumn"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_row_column_columns(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试按列批量更新。"""
    columns_data = [
        {"columnIndex": 0, "values": ["R1", "R2", "R3"]},
    ]

    def response_factory(request: dict) -> dict:
        assert request["elementId"] == "table-001"
        assert len(request["columns"]) == 1
        assert request["columns"][0]["values"] == ["R1", "R2", "R3"]

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_table_row_column_response(updated_count=1),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableRowColumn", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableRowColumn",
            params={
                "document_uri": client.document_uri,
                "elementId": "table-001",
                "columns": columns_data,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_row_column_both(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试同时提供行和列更新。"""

    def response_factory(request: dict) -> dict:
        assert "rows" in request
        assert "columns" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_table_row_column_response(updated_count=3),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableRowColumn", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableRowColumn",
            params={
                "document_uri": client.document_uri,
                "elementId": "table-001",
                "rows": [{"rowIndex": 0, "values": ["X", "Y"]}],
                "columns": [{"columnIndex": 1, "values": ["P", "Q"]}],
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["updatedCount"] == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_row_column_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Table not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableRowColumn", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableRowColumn",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
                "rows": [{"rowIndex": 0, "values": ["X"]}],
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
