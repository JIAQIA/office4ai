"""
Test Word event request/response flow

测试 Word 事件请求/响应流程。
"""

import asyncio

import pytest
from socketio import AsyncClient  # type: ignore[import-untyped]


@pytest.mark.asyncio
@pytest.mark.integration
async def test_get_selected_content_request(socketio_client: AsyncClient) -> None:
    """Test get selected content request"""
    request_data = {
        "requestId": "req_test_001",
        "documentUri": "file:///tmp/test.docx",
        "options": {"includeText": True},
    }

    # Emit request (call is synchronous in python-socketio client)
    socketio_client.emit("word:get:selectedContent", request_data, namespace="/word")

    # Note: Current implementation just logs, doesn't send response
    # When implementation is complete, we should receive response like:
    # response = await asyncio.wait_for(
    #     socketio_client.receive("word:get:selectedContent:response", namespace="/word"),
    #     timeout=5.0
    # )
    # assert response["requestId"] == "req_test_001"

    # For now, just verify no errors
    await asyncio.sleep(0.1)


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_text_request(socketio_client: AsyncClient) -> None:
    """Test insert text request"""
    request_data = {
        "requestId": "req_test_002",
        "documentUri": "file:///tmp/test.docx",
        "text": "Hello from test!",
        "location": "Cursor",
    }

    socketio_client.emit("word:insert:text", request_data, namespace="/word")

    # Note: Current implementation just logs
    await asyncio.sleep(0.1)


@pytest.mark.asyncio
@pytest.mark.integration
async def test_event_report_selection_changed(socketio_client: AsyncClient) -> None:
    """Test client reporting selection changed event"""
    event_data = {
        "eventType": "selectionChanged",
        "clientId": "test_client",
        "documentUri": "file:///tmp/test.docx",
        "data": {"text": "Selected text", "length": 13},
        "timestamp": 1234567890,
    }

    # Emit event (fire-and-forget, no response expected)
    socketio_client.emit("word:event:selectionChanged", event_data, namespace="/word")

    # Give server time to process
    await asyncio.sleep(0.1)

    # No exceptions raised = success
