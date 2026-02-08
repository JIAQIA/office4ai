"""Word MCP 工具集合 | Word MCP tools."""

from office4ai.a2c_smcp.tools.word.append_text import WordAppendTextTool
from office4ai.a2c_smcp.tools.word.get_selected_content import WordGetSelectedContentTool
from office4ai.a2c_smcp.tools.word.get_visible_content import WordGetVisibleContentTool
from office4ai.a2c_smcp.tools.word.insert_equation import WordInsertEquationTool
from office4ai.a2c_smcp.tools.word.insert_image import WordInsertImageTool
from office4ai.a2c_smcp.tools.word.insert_table import WordInsertTableTool
from office4ai.a2c_smcp.tools.word.insert_text import WordInsertTextTool
from office4ai.a2c_smcp.tools.word.insert_toc import WordInsertTOCTool
from office4ai.a2c_smcp.tools.word.replace_text import WordReplaceTextTool

__all__ = [
    "WordGetSelectedContentTool",
    "WordGetVisibleContentTool",
    "WordInsertTextTool",
    "WordAppendTextTool",
    "WordReplaceTextTool",
    "WordInsertImageTool",
    "WordInsertTableTool",
    "WordInsertEquationTool",
    "WordInsertTOCTool",
]
