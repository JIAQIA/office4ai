"""
Contract Tests for ppt:update:tableCell

测试 ppt:update:tableCell 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_cell_single(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试更新单个表格单元格。"""

    def response_factory(request: dict) -> dict:
        assert request["elementId"] == "table-001"
        assert len(request["cells"]) == 1
        cell = request["cells"][0]
        assert cell["rowIndex"] == 0
        assert cell["columnIndex"] == 1
        assert cell["text"] == "Updated"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_table_cell_response(updated_count=1),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableCell", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableCell",
            params={
                "document_uri": client.document_uri,
                "elementId": "table-001",
                "cells": [{"rowIndex": 0, "columnIndex": 1, "text": "Updated"}],
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["updatedCount"] == 1

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "ppt:update:tableCell"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_cell_multiple(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试批量更新多个单元格。"""
    cells = [
        {"rowIndex": 0, "columnIndex": 0, "text": "A1"},
        {"rowIndex": 0, "columnIndex": 1, "text": "B1"},
        {"rowIndex": 1, "columnIndex": 0, "text": "A2"},
    ]

    def response_factory(request: dict) -> dict:
        assert len(request["cells"]) == 3

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_table_cell_response(updated_count=3),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableCell", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableCell",
            params={
                "document_uri": client.document_uri,
                "elementId": "table-001",
                "cells": cells,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["updatedCount"] == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_cell_error(
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

    client.register_response("ppt:update:tableCell", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableCell",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
                "cells": [{"rowIndex": 0, "columnIndex": 0, "text": "X"}],
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
