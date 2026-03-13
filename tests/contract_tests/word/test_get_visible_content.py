"""
Contract Tests for word:get:visibleContent

测试 word:get:visibleContent 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


class TestLineBreakNormalization:
    """
    换行符规范化契约测试。

    验证 word:get:visibleContent 返回的文本符合 A2C 协议规范：
    - 必须使用 \\n (LF) 作为段落分隔符
    - 禁止使用 \\r (CR)
    - 禁止使用 \\r\\n (CRLF)

    参考：docs/a2c/line_break_normalization_spec.md
    """

    @pytest.mark.asyncio
    @pytest.mark.contract
    async def test_text_must_use_lf_not_cr(
        self,
        workspace: OfficeWorkspace,
        mock_word_client_factory,
        word_factory,
    ):
        """
        协议要求：text 字段必须使用 \\n 分隔段落，禁止使用 \\r。

        测试场景：
        - TypeScript 端应该已经将 Word 的 \\r\\r 转换为 \\n
        - Python 端验证返回的数据符合协议规范
        """
        # Arrange: 模拟 TypeScript 端正确转换后的数据
        normalized_text = "第一段\n第二段\n第三段\n"

        def response_with_normalized_text(request: dict) -> dict:
            return {
                "requestId": request["requestId"],
                "success": True,
                "data": word_factory.visible_content_response(text=normalized_text),
                "timestamp": int(asyncio.get_event_loop().time() * 1000),
            }

        client = mock_word_client_factory(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_test_word_client",
            document_uri="file:///tmp/contract_test.docx",
        )

        client.register_response("word:get:visibleContent", response_with_normalized_text)
        await client.connect()

        try:
            # Act
            action = OfficeAction(
                category="word",
                action_name="get:visibleContent",
                params={"document_uri": client.document_uri},
            )
            result = await workspace.execute(action)

            # Assert: 验证协议合规性
            assert result.success is True
            text = result.data["text"]

            # 禁止包含 \r
            assert "\r" not in text, (
                "协议违规：text 字段包含 \\r (CR)。"
                "TypeScript 端必须将换行符标准化为 \\n (LF)。"
                "参考：docs/a2c/line_break_normalization_spec.md"
            )

            # 验证段落使用 \n 分隔
            paragraphs = text.split("\n")
            assert len(paragraphs) == 4  # 三个段落 + 一个空结尾
            assert paragraphs[0] == "第一段"
            assert paragraphs[1] == "第二段"
            assert paragraphs[2] == "第三段"
            assert paragraphs[3] == ""  # 结尾的 \n

        finally:
            await client.disconnect()

    @pytest.mark.asyncio
    @pytest.mark.contract
    async def test_text_with_chinese_and_line_breaks(
        self,
        workspace: OfficeWorkspace,
        mock_word_client_factory,
        word_factory,
    ):
        """
        验证中文文本与换行符的正确组合。

        这是实际 E2E 测试中发现的问题场景。
        """
        # Arrange: 模拟实际场景（中文文本 + 图片标题）
        normalized_text = "文本与图片\n大标题\n"

        def response(request: dict) -> dict:
            return {
                "requestId": request["requestId"],
                "success": True,
                "data": word_factory.visible_content_response(text=normalized_text),
                "timestamp": int(asyncio.get_event_loop().time() * 1000),
            }

        client = mock_word_client_factory(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_test_word_client",
            document_uri="file:///tmp/contract_test.docx",
        )

        client.register_response("word:get:visibleContent", response)
        await client.connect()

        try:
            # Act
            action = OfficeAction(
                category="word",
                action_name="get:visibleContent",
                params={"document_uri": client.document_uri},
            )
            result = await workspace.execute(action)

            # Assert
            assert result.success is True
            text = result.data["text"]

            # 禁止 \r
            assert "\r" not in text, "协议违规：包含 \\r 字符"

            # 验证内容正确
            assert text == "文本与图片\n大标题\n"
            assert result.data["metadata"]["characterCount"] == 10

        finally:
            await client.disconnect()

    @pytest.mark.asyncio
    @pytest.mark.contract
    async def test_text_printable_in_terminal(
        self,
        workspace: OfficeWorkspace,
        mock_word_client_factory,
        word_factory,
    ):
        """
        终端可读性：文本应该能正确换行显示。

        如果包含 \\r，会导致终端显示混乱（覆盖行首）。
        """
        # Arrange
        normalized_text = "标题1\n内容1\n标题2\n内容2\n"

        def response(request: dict) -> dict:
            return {
                "requestId": request["requestId"],
                "success": True,
                "data": word_factory.visible_content_response(text=normalized_text),
                "timestamp": int(asyncio.get_event_loop().time() * 1000),
            }

        client = mock_word_client_factory(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_test_word_client",
            document_uri="file:///tmp/contract_test.docx",
        )

        client.register_response("word:get:visibleContent", response)
        await client.connect()

        try:
            # Act
            action = OfficeAction(
                category="word",
                action_name="get:visibleContent",
                params={"document_uri": client.document_uri},
            )
            result = await workspace.execute(action)

            # Assert: 验证可以在终端正确打印
            assert result.success is True
            text = result.data["text"]

            # 验证没有 \r 导致的覆盖问题
            assert "\r" not in text

            # 验证每个段落都是独立的
            lines = text.strip().split("\n")
            assert len(lines) == 4
            assert "标题1" in lines[0]
            assert "内容1" in lines[1]
            assert "标题2" in lines[2]
            assert "内容2" in lines[3]

        finally:
            await client.disconnect()

    @pytest.mark.asyncio
    @pytest.mark.contract
    async def test_empty_paragraphs_preserved(
        self,
        workspace: OfficeWorkspace,
        mock_word_client_factory,
        word_factory,
    ):
        """
        空段落应该被保留。

        Word 文档中的空段落应该转换为 \\n，而不是被删除。
        """
        # Arrange: 包含空段落的文本
        normalized_text = "第一段\n\n第三段\n"  # 第二段是空的

        def response(request: dict) -> dict:
            return {
                "requestId": request["requestId"],
                "success": True,
                "data": word_factory.visible_content_response(text=normalized_text),
                "timestamp": int(asyncio.get_event_loop().time() * 1000),
            }

        client = mock_word_client_factory(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_test_word_client",
            document_uri="file:///tmp/contract_test.docx",
        )

        client.register_response("word:get:visibleContent", response)
        await client.connect()

        try:
            # Act
            action = OfficeAction(
                category="word",
                action_name="get:visibleContent",
                params={"document_uri": client.document_uri},
            )
            result = await workspace.execute(action)

            # Assert
            assert result.success is True
            text = result.data["text"]

            # 验证空段落被保留（两个连续的 \n）
            assert text == "第一段\n\n第三段\n"
            paragraphs = text.split("\n")
            assert paragraphs[0] == "第一段"
            assert paragraphs[1] == ""  # 空段落
            assert paragraphs[2] == "第三段"

        finally:
            await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_visible_content_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试成功获取可见内容的完整流程。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应
    3. 连接到 Workspace
    4. Workspace 发送 word:get:visibleContent 命令
    5. 验证返回的响应数据符合协议
    """
    # Arrange: 创建并配置 Mock 客户端
    expected_text = "Test visible content from contract test"

    def visible_content_response(request: dict) -> dict:
        """动态响应工厂，验证请求数据并返回响应"""
        # 验证请求数据
        assert "requestId" in request
        assert "documentUri" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.visible_content_response(text=expected_text),
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
    client.register_response("word:get:visibleContent", visible_content_response)

    # 连接到服务器
    await client.connect()

    try:
        # Act: 执行动作
        action = OfficeAction(
            category="word",
            action_name="get:visibleContent",
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
        assert event_name == "word:get:visibleContent"
        assert "requestId" in event_data
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_visible_content_empty_document_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """
    测试空文档的错误处理。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册错误响应（文档未找到或为空）
    3. 连接到 Workspace
    4. Workspace 发送 word:get:visibleContent 命令
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
                "message": "Document not found or empty",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:visibleContent", error_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:visibleContent",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is False
        assert "3001" in str(result.error) or "not found" in str(result.error).lower()
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_visible_content_with_options(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试带选项的获取可见内容。

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
        "maxTextLength": 1000,
    }

    def response_with_options_validation(request: dict) -> dict:
        """验证选项的响应工厂"""
        # 验证选项正确传递
        assert "options" in request
        actual_options = request["options"]

        # 验证我们关心的字段
        for key, value in expected_options.items():
            assert actual_options.get(key) == value, f"Expected {key}={value}, got {actual_options.get(key)}"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.visible_content_response(
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

    client.register_response("word:get:visibleContent", response_with_options_validation)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:visibleContent",
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
async def test_get_visible_content_with_complex_elements(
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
        "word:get:visibleContent",
        {
            "requestId": "test_req_001",
            "success": True,
            "data": word_factory.visible_content_response(
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
            action_name="get:visibleContent",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["text"] == "Complex content with image and table"

        # 验证元素列表
        assert "elements" in result.data
        assert len(result.data["elements"]) >= 2  # 至少包含文本和其他元素
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_visible_content_with_text_length_limit(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试文本长度限制功能。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应，验证 maxTextLength 选项
    3. 连接到 Workspace
    4. Workspace 发送带长度限制的命令
    5. 验证返回的文本长度符合限制
    """
    # Arrange
    max_length = 100

    def response_with_length_validation(request: dict) -> dict:
        """验证文本长度限制的响应工厂"""
        options = request.get("options", {})
        assert options.get("maxTextLength") == max_length

        # 返回的文本应该符合长度限制
        full_text = "A" * 200  # 200字符
        truncated_text = full_text[:max_length]

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.visible_content_response(text=truncated_text),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:visibleContent", response_with_length_validation)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:visibleContent",
            params={
                "document_uri": client.document_uri,
                "options": {"maxTextLength": max_length},
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert len(result.data["text"]) <= max_length
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_visible_content_timeout(
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

    # Arrange: 临时缩短超时时间，避免测试等待过久
    original_timeout = workspace.config.request_timeout
    workspace.config.request_timeout = 1000  # 1 秒

    # 注册延迟响应
    async def slow_response(request: dict) -> dict:
        """延迟响应工厂"""
        await asyncio.sleep(2)  # 超过测试超时时间（1秒）
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.visible_content_response(text="Delayed response"),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:visibleContent", slow_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:visibleContent",
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
        workspace.config.request_timeout = original_timeout
        await client.disconnect()
