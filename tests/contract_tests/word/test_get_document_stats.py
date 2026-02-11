"""
Contract Tests for word:get:documentStats

测试 word:get:documentStats 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_document_stats_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试成功获取文档统计信息的完整流程。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应
    3. 连接到 Workspace
    4. Workspace 发送 word:get:documentStats 命令
    5. 验证返回的响应数据符合协议
    """
    # Arrange: 创建并配置 Mock 客户端
    expected_stats = {
        "wordCount": 1500,
        "characterCount": 7500,
        "paragraphCount": 30,
    }

    def document_stats_response(request: dict) -> dict:
        """动态响应工厂，验证请求数据并返回响应"""
        # 验证请求数据
        assert "requestId" in request
        assert "documentUri" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.document_stats_response(
                word_count=expected_stats["wordCount"],
                character_count=expected_stats["characterCount"],
                paragraph_count=expected_stats["paragraphCount"],
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    # 创建客户端
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    # 在连接之前注册响应
    client.register_response("word:get:documentStats", document_stats_response)

    # 连接到服务器
    await client.connect()

    try:
        # Act: 执行动作
        action = OfficeAction(
            category="word",
            action_name="get:documentStats",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert: 验证结果
        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["wordCount"] == expected_stats["wordCount"]
        assert result.data["characterCount"] == expected_stats["characterCount"]
        assert result.data["paragraphCount"] == expected_stats["paragraphCount"]

        # 验证 Mock 客户端接收到了请求
        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:get:documentStats"
        assert "requestId" in event_data
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_document_stats_empty_document(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试空文档的统计信息。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册空文档响应
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证空文档的正确计数
    """
    # Arrange: 创建并配置 Mock 客户端
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/empty_test.docx",
    )

    client.register_static_response(
        "word:get:documentStats",
        {
            "requestId": "test_req_001",
            "success": True,
            "data": word_factory.document_stats_response(
                word_count=0,
                character_count=0,
                paragraph_count=0,
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        },
    )

    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:documentStats",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["wordCount"] == 0
        assert result.data["characterCount"] == 0
        assert result.data["paragraphCount"] == 0
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_document_stats_large_document(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试大文档的统计信息。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册大文档响应
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证大文档的正确计数
    """
    # Arrange: 创建并配置 Mock 客户端
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/large_test.docx",
    )

    client.register_static_response(
        "word:get:documentStats",
        {
            "requestId": "test_req_002",
            "success": True,
            "data": word_factory.document_stats_response(
                word_count=100000,
                character_count=500000,
                paragraph_count=1000,
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        },
    )

    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:documentStats",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["wordCount"] == 100000
        assert result.data["characterCount"] == 500000
        assert result.data["paragraphCount"] == 1000
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_document_stats_document_not_found(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """
    测试文档未找到的错误处理。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册错误响应（文档未找到）
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证错误信息正确返回
    """

    # Arrange: 创建并配置 Mock 客户端
    def error_response(request: dict) -> dict:
        """错误响应工厂"""
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3001",
                "message": "Document not found",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/nonexistent.docx",
    )

    client.register_response("word:get:documentStats", error_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:documentStats",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is False
        assert "3001" in str(result.error) or "not found" in str(result.error).lower()
    finally:
        await client.disconnect()
