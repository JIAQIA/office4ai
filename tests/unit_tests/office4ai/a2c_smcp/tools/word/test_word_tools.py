"""
Word MCP Tools 单元测试 | Word MCP Tools unit tests

测试策略:
- Mock OfficeWorkspace.execute() 的返回值
- 验证 OfficeAction 构建 (category, event_name, params)
- 验证输入校验 (缺少参数、参数类型错误)
- 验证 format_result hook (获取类 vs 操作类)
"""

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.tools.word import (
    WordAppendTextTool,
    WordGetSelectedContentTool,
    WordGetVisibleContentTool,
    WordInsertEquationTool,
    WordInsertImageTool,
    WordInsertTableTool,
    WordInsertTextTool,
    WordInsertTOCTool,
    WordReplaceTextTool,
)
from office4ai.environment.workspace.base import OfficeObs


@pytest.fixture
def mock_workspace():
    """创建 mock OfficeWorkspace"""
    workspace = MagicMock()
    workspace.execute = AsyncMock()
    return workspace


# ============================================================================
# Tool Metadata Tests
# ============================================================================


class TestToolMetadata:
    """测试所有工具的元数据声明"""

    TOOL_SPECS = [
        (WordGetSelectedContentTool, "word_get_selected_content", "word", "get:selectedContent"),
        (WordGetVisibleContentTool, "word_get_visible_content", "word", "get:visibleContent"),
        (WordInsertTextTool, "word_insert_text", "word", "insert:text"),
        (WordAppendTextTool, "word_append_text", "word", "append:text"),
        (WordReplaceTextTool, "word_replace_text", "word", "replace:text"),
        (WordInsertImageTool, "word_insert_image", "word", "insert:image"),
        (WordInsertTableTool, "word_insert_table", "word", "insert:table"),
        (WordInsertEquationTool, "word_insert_equation", "word", "insert:equation"),
        (WordInsertTOCTool, "word_insert_toc", "word", "insert:toc"),
    ]

    @pytest.mark.parametrize("tool_cls,expected_name,expected_category,expected_event", TOOL_SPECS)
    def test_tool_metadata(self, mock_workspace, tool_cls, expected_name, expected_category, expected_event):
        """验证工具元数据 | Verify tool metadata"""
        tool = tool_cls(mock_workspace)
        assert tool.name == expected_name
        assert tool.category == expected_category
        assert tool.event_name == expected_event
        assert isinstance(tool.description, str)
        assert len(tool.description) > 0
        assert isinstance(tool.input_schema, dict)
        assert "properties" in tool.input_schema

    @pytest.mark.parametrize("tool_cls,expected_name,expected_category,expected_event", TOOL_SPECS)
    def test_input_schema_has_document_uri(
        self, mock_workspace, tool_cls, expected_name, expected_category, expected_event
    ):
        """验证所有工具的 input_schema 都包含 document_uri"""
        tool = tool_cls(mock_workspace)
        schema = tool.input_schema
        assert "document_uri" in schema["properties"]


# ============================================================================
# Execute Flow Tests
# ============================================================================


class TestExecuteFlow:
    """测试通用执行流程"""

    @pytest.mark.asyncio
    async def test_insert_text_builds_correct_action(self, mock_workspace):
        """验证 insert_text 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Hello World",
            }
        )

        mock_workspace.execute.assert_called_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:text"
        assert action.params["document_uri"] == "file:///test.docx"
        assert action.params["text"] == "Hello World"
        assert action.params["location"] == "Cursor"  # default value

    @pytest.mark.asyncio
    async def test_insert_text_with_format(self, mock_workspace):
        """验证带格式的文本插入"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Bold Text",
                "format": {"bold": True, "font_size": 14},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["text"] == "Bold Text"
        assert action.params["format"]["bold"] is True
        assert action.params["format"]["font_size"] == 14

    @pytest.mark.asyncio
    async def test_append_text_builds_correct_action(self, mock_workspace):
        """验证 append_text 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordAppendTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Appended",
                "location": "Start",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "append:text"
        assert action.params["location"] == "Start"

    @pytest.mark.asyncio
    async def test_replace_text_builds_correct_action(self, mock_workspace):
        """验证 replace_text 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"replaceCount": 3})

        tool = WordReplaceTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "foo",
                "replace_text": "bar",
                "options": {"match_case": True, "replace_all": True},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "replace:text"
        assert action.params["search_text"] == "foo"
        assert action.params["replace_text"] == "bar"
        assert action.params["options"]["match_case"] is True
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_insert_image_builds_correct_action(self, mock_workspace):
        """验证 insert_image 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertImageTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "image": {"base64": "iVBORw0KGgo=", "width": 100, "height": 100},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:image"
        assert action.params["image"]["base64"] == "iVBORw0KGgo="

    @pytest.mark.asyncio
    async def test_insert_table_builds_correct_action(self, mock_workspace):
        """验证 insert_table 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTableTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "options": {"rows": 3, "columns": 2, "data": [["A", "B"], ["C", "D"], ["E", "F"]]},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:table"
        assert action.params["options"]["rows"] == 3
        assert action.params["options"]["columns"] == 2

    @pytest.mark.asyncio
    async def test_insert_equation_builds_correct_action(self, mock_workspace):
        """验证 insert_equation 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertEquationTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "latex": "E = mc^2",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:equation"
        assert action.params["latex"] == "E = mc^2"

    @pytest.mark.asyncio
    async def test_insert_toc_builds_correct_action(self, mock_workspace):
        """验证 insert_toc 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTOCTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "options": {"max_level": 2},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:toc"
        assert action.params["options"]["max_level"] == 2


# ============================================================================
# Input Validation Tests
# ============================================================================


class TestInputValidation:
    """测试输入验证"""

    @pytest.mark.asyncio
    async def test_missing_document_uri(self, mock_workspace):
        """测试缺少 document_uri"""
        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute({"text": "Hello"})

        assert result["success"] is False
        assert "error" in result
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_missing_required_field(self, mock_workspace):
        """测试缺少必要字段"""
        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is False
        assert "error" in result
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_replace_text_search_too_long(self, mock_workspace):
        """测试搜索文本超过 255 字符"""
        tool = WordReplaceTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "a" * 256,
                "replace_text": "b",
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()


# ============================================================================
# Format Result Tests
# ============================================================================


class TestFormatResult:
    """测试 format_result hook"""

    @pytest.mark.asyncio
    async def test_get_selected_content_format(self, mock_workspace):
        """获取类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"text": "Selected text content", "elements": []},
        )

        tool = WordGetSelectedContentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
            }
        )

        assert result["success"] is True
        assert result["content"] == "Selected text content"
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_visible_content_format(self, mock_workspace):
        """获取类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"text": "Visible content here", "elements": []},
        )

        tool = WordGetVisibleContentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
            }
        )

        assert result["success"] is True
        assert result["content"] == "Visible content here"

    @pytest.mark.asyncio
    async def test_operation_tool_format(self, mock_workspace):
        """操作类工具返回标准 JSON (data 字段)"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"inserted": True},
        )

        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Hello",
            }
        )

        assert result["success"] is True
        assert result["data"] == {"inserted": True}

    @pytest.mark.asyncio
    async def test_error_format(self, mock_workspace):
        """错误返回统一格式"""
        mock_workspace.execute.return_value = OfficeObs(
            success=False,
            data={},
            error="Document not connected: file:///test.docx",
        )

        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Hello",
            }
        )

        assert result["success"] is False
        assert "Document not connected" in result["error"]

    @pytest.mark.asyncio
    async def test_workspace_exception(self, mock_workspace):
        """workspace 抛出异常时的处理"""
        mock_workspace.execute.side_effect = TimeoutError("Operation timed out")

        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Hello",
            }
        )

        assert result["success"] is False
        assert "timed out" in result["error"]
