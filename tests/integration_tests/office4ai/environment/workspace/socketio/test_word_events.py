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

    # Emit request
    await socketio_client.emit("word:get:selectedContent", request_data, namespace="/word")

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

    await socketio_client.emit("word:insert:text", request_data, namespace="/word")

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
    await socketio_client.emit("word:event:selectionChanged", event_data, namespace="/word")

    # Give server time to process
    await asyncio.sleep(0.1)

    # No exceptions raised = success


@pytest.mark.asyncio
@pytest.mark.integration
async def test_replace_text_request(socketio_client: AsyncClient) -> None:
    """Test replace text request"""
    request_data = {
        "requestId": "req_test_replace_001",
        "documentUri": "file:///tmp/test.docx",
        "searchText": "old text",
        "replaceText": "new text",
        "options": {
            "matchCase": False,
            "matchWholeWord": False,
            "replaceAll": True,
        },
    }

    await socketio_client.emit("word:replace:text", request_data, namespace="/word")

    # Note: Current implementation just logs
    await asyncio.sleep(0.1)


@pytest.mark.asyncio
@pytest.mark.integration
async def test_replace_text_with_match_case(socketio_client: AsyncClient) -> None:
    """Test replace text with case sensitivity"""
    request_data = {
        "requestId": "req_test_replace_002",
        "documentUri": "file:///tmp/test.docx",
        "searchText": "Hello",
        "replaceText": "Hi",
        "options": {
            "matchCase": True,
        },
    }

    await socketio_client.emit("word:replace:text", request_data, namespace="/word")

    await asyncio.sleep(0.1)


@pytest.mark.asyncio
@pytest.mark.integration
async def test_replace_text_with_whole_word(socketio_client: AsyncClient) -> None:
    """Test replace text with whole word matching"""
    request_data = {
        "requestId": "req_test_replace_003",
        "documentUri": "file:///tmp/test.docx",
        "searchText": "test",
        "replaceText": "exam",
        "options": {
            "matchWholeWord": True,
            "replaceAll": True,
        },
    }

    await socketio_client.emit("word:replace:text", request_data, namespace="/word")

    await asyncio.sleep(0.1)


@pytest.mark.asyncio
@pytest.mark.integration
async def test_replace_text_empty_validation(socketio_client: AsyncClient) -> None:
    """Test replace text with empty search/replace strings (should log warning)"""
    request_data = {
        "requestId": "req_test_replace_004",
        "documentUri": "file:///tmp/test.docx",
        "searchText": "",  # Empty - should trigger warning
        "replaceText": "",  # Empty - should trigger warning
        "options": {},
    }

    await socketio_client.emit("word:replace:text", request_data, namespace="/word")

    await asyncio.sleep(0.1)
