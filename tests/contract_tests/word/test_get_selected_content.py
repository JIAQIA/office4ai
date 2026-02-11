"""
Contract Tests for word:get:selectedContent

测试 word:get:selectedContent 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selected_content_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试成功获取选中内容的完整流程。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应
    3. 连接到 Workspace
    4. Workspace 发送 word:get:selectedContent 命令
    5. 验证返回的响应数据符合协议
    """
    # Arrange: 创建并配置 Mock 客户端
    expected_text = "Test selected content from contract test"

    def selected_content_response(request: dict) -> dict:
        """动态响应工厂，验证请求数据并返回响应"""
        # 验证请求数据
        assert "requestId" in request
        assert "documentUri" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.selected_content_response(text=expected_text),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
            "duration": 50,
        }

    # 创建客户端
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    # 在连接之前注册响应
    client.register_response("word:get:selectedContent", selected_content_response)

    # 连接到服务器
    await client.connect()

    try:
        # Act: 执行动作
        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert: 验证结果
        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["text"] == expected_text
        assert "metadata" in result.data
        assert result.data["metadata"]["characterCount"] == len(expected_text)
        assert "elements" in result.data

        # 验证 Mock 客户端接收到了请求
        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:get:selectedContent"
        assert "requestId" in event_data
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selected_content_empty_selection_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """
    测试选区为空的错误处理。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册错误响应（选区为空）
    3. 连接到 Workspace
    4. Workspace 发送 word:get:selectedContent 命令
    5. 验证错误信息正确返回
    """

    # Arrange: 创建并配置 Mock 客户端
    def error_response(request: dict) -> dict:
        """错误响应工厂"""
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3002",
                "message": "Selection is empty",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:selectedContent", error_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is False
        assert "3002" in str(result.error) or "empty" in str(result.error).lower()
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selected_content_with_options(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试带选项的获取选中内容。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应
    3. 连接到 Workspace
    4. Workspace 发送带选项的命令
    5. 验证选项正确传递到 Mock 客户端
    """
    # Arrange
    expected_options = {
        "includeText": True,
        "includeImages": True,
        "includeTables": False,
        "detailedMetadata": True,
    }

    def response_with_options_validation(request: dict) -> dict:
        """验证选项的响应工厂"""
        # 验证选项正确传递（检查关键字段，而不是完全相等）
        assert "options" in request
        actual_options = request["options"]

        # 验证我们关心的字段
        for key, value in expected_options.items():
            assert actual_options.get(key) == value, f"Expected {key}={value}, got {actual_options.get(key)}"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.selected_content_response(
                text="Content with options",
                include_images=True,
                include_tables=False,
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:selectedContent", response_with_options_validation)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={
                "document_uri": client.document_uri,
                "options": expected_options,
            },
        )
        result = await workspace.execute(action)

        # Assert
        if not result.success:
            pytest.fail(
                f"Expected success, got error: '{result.error}', data: {result.data}, metadata: {result.metadata}"
            )
        assert result.success is True
        assert result.data["text"] == "Content with options"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selected_content_with_complex_elements(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试包含复杂元素（图片、表格）的响应。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册包含图片和表格的响应
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证复杂元素正确返回
    """
    # Arrange
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_static_response(
        "word:get:selectedContent",
        {
            "requestId": "test_req_001",
            "success": True,
            "data": word_factory.selected_content_response(
                text="Complex content with image and table",
                include_images=True,
                include_tables=True,
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        },
    )

    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["text"] == "Complex content with image and table"

        # 验证元数据
        assert result.data["metadata"]["imageCount"] == 1
        assert result.data["metadata"]["tableCount"] == 1

        # 验证元素列表
        assert len(result.data["elements"]) == 3  # paragraph + image + table
        element_types = [elem["type"] for elem in result.data["elements"]]
        assert "paragraph" in element_types
        assert "image" in element_types
        assert "table" in element_types
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selected_content_timeout(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试请求超时场景。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册延迟响应
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证超时错误
    """

    # Arrange: 注册延迟响应
    async def slow_response(request: dict) -> dict:
        """延迟响应工厂"""
        await asyncio.sleep(15)  # 超过默认超时时间（10秒）
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.selected_content_response(text="Delayed response"),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:selectedContent", slow_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        if result.success:
            pytest.fail(f"Expected timeout error, but got success: {result.data}")
        # 验证有错误信息（即使不是标准的 timeout 消息）
        assert result.error is not None or result.metadata.get("error"), (
            f"Expected error, got: error='{result.error}', metadata={result.metadata}"
        )
    finally:
        await client.disconnect()
