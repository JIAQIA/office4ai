# filename: schema.py
# @Time    : 2024/4/16 18:57
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import datetime
from enum import Enum
from typing import Annotated, Any, Literal

from pydantic import BaseModel, BeforeValidator, Field, Json

ACTION_CATEGORY_MAP: dict[int, str] = {
    0: "terminal",
    1: "workspace",
}

TEXT_DOCUMENT_ACTIONS: set[str] = {  # text_document
    "open_file",
    "apply_edit",
    "save_file",
    "close_file",
    "read_file",
    "get_file_symbols",
    "find_in_file",
    "replace_in_file",
    "get_definition_and_implementation",
    "hover",
    "insert_cursor",
    "delete_cursor",
    "clear_cursors",
}

WORKSPACE_ACTIONS: set[str] = {
    # workspace
    "find_in_workspace",
    "replace_in_workspace",
    "create_file",
    "create_files",
    "delete_file",
    "delete_files",
    "rename_file",
}

LSP_ACTIONS: set[str] = {
    # lsp
    "restart_lsp",
}

TERMINAL_ACTIONS: set[str] = set()

ACTIONS: set[str] = TEXT_DOCUMENT_ACTIONS | WORKSPACE_ACTIONS | LSP_ACTIONS | TERMINAL_ACTIONS

ACTION_NAME_MAP: dict[int, str] = dict(enumerate(ACTIONS))


class IDEObs(BaseModel):
    """
    The observation type of the IDE environment.
    """

    created_at: str = Field(
        default_factory=lambda: datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime(
            "%Y-%m-%d %H:%M:%S",
        ),
        title="创建时间",
        description="当前观察的创建时间，使用北京时间（东八区）",
        pattern=r"^(19|20)\d\d-(0[1-9]|1[0-2])-(0[1-9]|[12][0-9]|3[01]) (0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]):"
        "([0-5][0-9])$",
    )
    obs: str | None = Field(
        default=None,
        title="观察内容",
        description="当前观察的内容",
    )
    original_result: Any = Field(
        default=None,
        title="原始结果",
        description="IDE执行返回的原始结果",
    )


class IDEAction(BaseModel):
    """
    The action type of the IDE environment.
    """

    category: Annotated[
        Literal["terminal", "workspace"],
        BeforeValidator(lambda s: s if isinstance(s, str) else ACTION_CATEGORY_MAP[s]),
    ] = Field(title="动作类别", description="动作的类别，是终端动作还是工作区动作", frozen=True)

    action_name: Annotated[str, BeforeValidator(lambda s: s if isinstance(s, str) else ACTION_NAME_MAP[s])] = Field(
        title="动作名称",
        description="一般是指命令行命令的名称，比如grep，或者可执行的工具名称，比如insert",
        frozen=True,
    )

    action_args: dict | Json[Any] | list[str] | str | None = Field(
        default=None,
        title="动作参数",
        union_mode="left_to_right",
        description="动作的参数，比如grep的参数，或者insert的参数，参数可以以key-value组成的dict传入",
        frozen=True,
    )


class LanguageId(Enum):
    """
    The language identifier of the IDE environment.
    """

    python = "python"
