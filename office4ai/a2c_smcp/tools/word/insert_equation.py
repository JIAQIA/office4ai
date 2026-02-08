"""word_insert_equation MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import EquationOptions


class WordInsertEquationInput(BaseModel):
    """MCP 输入模型: 插入公式"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    latex: str = Field(..., description="LaTeX equation string (e.g. 'E = mc^2')")
    options: EquationOptions | None = Field(None, description="Equation options (inline display)")


class WordInsertEquationTool(BaseTool):
    """在 Word 文档中插入 LaTeX 公式"""

    @property
    def name(self) -> str:
        return "word_insert_equation"

    @property
    def description(self) -> str:
        return (
            "Insert a LaTeX equation into a Word document. "
            "Provide the equation as a LaTeX string. "
            "By default, the equation is inserted inline."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertEquationInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:equation"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertEquationInput
