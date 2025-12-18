# filename: workspace.py
# @Time    : 2024/4/30 17:08
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import datetime
import json
import os
import subprocess
from collections.abc import Callable, Sequence
from typing import Any, SupportsFloat, cast

from pydantic import AnyUrl, ValidationError

from ide4ai.dtos.base_protocol import LSPResponseMessage
from ide4ai.dtos.diagnostics import DocumentDiagnosticReport
from ide4ai.dtos.workspace_edit import LSPWorkspaceEdit
from ide4ai.environment.workspace.base import BaseWorkspace
from ide4ai.environment.workspace.model import TextModel
from ide4ai.environment.workspace.schema import Position, Range, SearchResult, SingleEditOperation, TextEdit
from ide4ai.exceptions import IDEExecutionError
from ide4ai.python_ide.const import DEFAULT_CAPABILITY, DEFAULT_SYMBOL_VALUE_SET
from ide4ai.schema import LSP_ACTIONS, TEXT_DOCUMENT_ACTIONS, WORKSPACE_ACTIONS, IDEAction, IDEObs, LanguageId
from ide4ai.utils import render_symbols


def default_python_header_generator(workspace: BaseWorkspace, file_path: str) -> str:
    """
    默认的Python文件头生成器

    Args:
        workspace (BaseWorkspace): 工作环境
        file_path (str): 文件路径

    Returns:
        str: 文件头
    """
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
    return (
        f"# -*- coding: utf-8 -*-\n"
        f"# filename : {os.path.basename(file_path)}\n"
        f"# @Time    : {now.strftime('%Y/%m/%d %H:%M')}\n"
        f"# @Author  : TuringFocus\n"
        f"# @Email   : support@turingfocus.com\n"
        f"# @Software: {workspace.project_name}\n"
    )


class PyWorkspace(BaseWorkspace):
    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        if self.header_generators is None:
            self.header_generators: dict[str, Callable[[BaseWorkspace, str], str]] = {
                ".py": default_python_header_generator,
            }
        # 如果没有提供shortcut_commands，尝试自动检测Makefile | Auto-detect Makefile if no shortcut_commands provided
        if self.shortcut_commands is None:
            from ide4ai.utils import detect_makefile_commands

            self.shortcut_commands = detect_makefile_commands(self.root_dir)

    @staticmethod
    def _format_diagnostics(diagnostics: DocumentDiagnosticReport | None) -> str:
        """
        格式化诊断信息为可读字符串 / Format diagnostics to readable string

        Args:
            diagnostics: 诊断结果 / Diagnostics result

        Returns:
            str: 格式化后的诊断信息 / Formatted diagnostics string
        """
        if not diagnostics:
            return ""

        result = "\n\n诊断信息 / Diagnostics:\n"

        # 处理文档诊断报告 / Handle document diagnostic report
        if hasattr(diagnostics, "items"):
            # Full report with diagnostic items
            if not diagnostics.items:
                result += "  无诊断问题 / No diagnostic issues\n"
            else:
                for idx, diagnostic in enumerate(diagnostics.items, 1):
                    severity = diagnostic.severity
                    severity_str = (
                        {1: "错误/Error", 2: "警告/Warning", 3: "信息/Info", 4: "提示/Hint"}.get(
                            severity,
                            "未知/Unknown",
                        )
                        if isinstance(severity, int)
                        else "未知/Unknown"
                    )
                    message = getattr(diagnostic, "message", "")
                    range_info = getattr(diagnostic, "range", None)
                    if range_info:
                        start = getattr(range_info, "start", None)
                        line = getattr(start, "line", "?") if start else "?"
                        result += f"  [{idx}] {severity_str} (行{line}): {message}\n"
                    else:
                        result += f"  [{idx}] {severity_str}: {message}\n"
        elif hasattr(diagnostics, "kind") and diagnostics.kind == "unchanged":
            result += "  诊断信息未变化 / Diagnostics unchanged\n"
        else:
            result += f"  诊断结果: {diagnostics}\n"

        return result

    def _launch_lsp(self) -> subprocess.Popen[bytes]:
        """
        启动 Pyright 语言服务器 / Launch Pyright language server

        注意启动时需要使用Bytes模式，而不是Str模式，即text设置为False。因为LSP协议长度计算是按bytes来计算的。

        Returns:
            subprocess.Popen[bytes]: Pyright 语言服务器进程 | Pyright language server process
        """
        # 启动 Pyright 语言服务器
        process = subprocess.Popen(
            ["pyright-langserver", "--stdio"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            text=False,
        )
        return process

    def _initial_lsp(self) -> None:
        """
        初始化 LSP 服务 / Initialize LSP service

        Returns:

        """
        msg_id = self.get_lsp_msg_id()
        res = self.send_lsp_msg(
            "initialize",
            {
                "processId": None,
                "workspaceFolders": [
                    {
                        "uri": f"file://{self.root_dir}",
                        "name": self.project_name,
                    },
                ],
                "initializationOptions": {
                    "disablePullDiagnostics": False,  # 启用 Pull Diagnostics / Enable Pull Diagnostics
                },
                "capabilities": DEFAULT_CAPABILITY,
            },
            message_id=msg_id,
        )
        if res:
            try:
                res_json = LSPResponseMessage.model_validate(json.loads(res))
            except json.JSONDecodeError as e:  # pragma: no cover
                raise ValueError(f"初始化LSP服务失败，返回结果无法解析为json: {res}") from e  # pragma: no cover
            if res_json.error:
                raise ValueError(f"初始化LSP服务失败: {res_json.error}")  # pragma: no cover
            self.send_lsp_msg("initialized")

    def construct_action(self, action: dict) -> IDEAction:
        """
        构建 IDEAction 对象

        Args:
            action (dict): 动作字典 | Action dictionary

        Returns:
            IDEAction: IDEAction 对象 | IDEAction object

        Raises:
            ValueError: 如果动作类别为 "terminal" | If the action category is "terminal"
            ValueError: 如果动作不在支持的动作集合中 | If the action is not in the supported action set
        """
        ide_action = IDEAction.model_validate(action)
        match ide_action.category:
            case "terminal":
                raise ValueError("Workspace不支持terminal的动作类别")
            case "workspace":
                if ide_action.action_name not in WORKSPACE_ACTIONS | TEXT_DOCUMENT_ACTIONS | LSP_ACTIONS:
                    raise ValueError(f"Workspace不支持 {ide_action.action_name} 动作")
                return ide_action
            case _:
                raise ValueError(f"不支持的动作类别 {ide_action.category}")  # pragma: no cover

    def step(self, action: dict) -> tuple[dict, SupportsFloat, bool, bool, dict[str, Any]]:
        """
        执行一个动作

        观察返回：
        1. OpenFile: 返回打开文件的内容
        2. ApplyEdit: 返回编辑的变更记录

        奖励机制：
        1. OpenFile: 成功打印返回100，打开失败返回0
        2. ApplyEdit: 变更成功返回100，失败返回0  TODO 后续需要有更细致的评估方法来评估编辑的质量

        Args:
            action (dict): 动作字典 | Action dictionary

        Returns:
            tuple[dict, SupportsFloat, bool, bool, dict[str, Any]]: 观察、奖励、是否结束、是否成功、额外信息 |
                Observation, Reward, Done, Success, Extra info
        """
        self._assert_not_closed()
        ide_action = self.construct_action(action)
        match ide_action.action_name:
            case "open_file":
                try:
                    if isinstance(ide_action.action_args, dict):
                        text_model = self.open_file(**ide_action.action_args)
                    elif isinstance(ide_action.action_args, str):
                        text_model = self.open_file(uri=ide_action.action_args)
                    else:
                        raise ValueError("open_file 动作参数错误")
                    file_content = (
                        text_model.get_simple_view() if self._enable_simple_view_mode else text_model.get_view()
                    )
                    return (
                        IDEObs(obs=file_content).model_dump(),
                        100,
                        True,
                        True,
                        {},
                    )
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "apply_edit":
                try:
                    if isinstance(ide_action.action_args, dict):
                        res, diagnostics = self.apply_edit(**ide_action.action_args)
                        # 成功编辑后，读取编辑位置附近的代码并返回，给LLM一个反馈
                        content_after_edit: list[str] = []
                        content_ranges: list[Range] = []
                        for r in ide_action.action_args.get("edits", []):
                            if isinstance(r, dict):
                                edit = SingleEditOperation.model_validate(r)
                                length_of_new_text = len(edit.text.splitlines()) if edit.text else 0
                                range_start_line = edit.range.start_position.line
                                range_end_line = edit.range.end_position.line
                                content_ranges.append(
                                    Range(
                                        start_position=Position(max(1, range_start_line - 3), 1),
                                        end_position=Position(range_end_line + length_of_new_text + 3, 1),
                                    ),
                                )
                        # 对content_ranges进行合并。有交集的合并，无交集独立
                        content_ranges.sort(key=lambda x: x.start_position.line)
                        merged_ranges: list[Range] = []
                        for r in content_ranges:
                            if not merged_ranges:
                                merged_ranges.append(r)
                            else:
                                last_range = merged_ranges[-1]
                                if last_range & r:
                                    last_range |= r
                                else:
                                    merged_ranges.append(r)
                        if merged_ranges:
                            # TODO 这里的展示效果不好，原因是model.get_view不允许传入多个Range，仅允许单个Range，
                            #  这导致文本渲染会有大量重复。mode.get_view需要支持多Range模式
                            content_after_edit.append(
                                self.read_file(
                                    uri=ide_action.action_args["uri"],
                                    with_line_num=True,
                                    code_range=merged_ranges[-1],
                                ),
                            )
                        apply_result: str = (
                            "编辑成功。如果有回滚需求，可以按下面的回滚操作执行。" + "\n".join([repr(e) for e in res])
                            if res
                            else ""
                        )
                        if content_after_edit:
                            apply_result += "\n编辑后的代码如下（仅返回编辑位置附近的代码。如果想看全部，可以使用read_file工具查看）:\n"
                            apply_result += "\n".join(content_after_edit)

                        # 添加诊断信息到返回结果 / Add diagnostics to result
                        if diagnostics:
                            apply_result += self._format_diagnostics(diagnostics)

                        return IDEObs(obs=apply_result, original_result=res).model_dump(), 100, True, True, {}
                    else:
                        raise ValueError("apply_edit 动作参数错误")
                except IDEExecutionError as e:
                    raise e
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "save_file":
                try:
                    if isinstance(ide_action.action_args, dict):
                        self.save_file(**ide_action.action_args)
                    elif isinstance(ide_action.action_args, str):
                        self.save_file(uri=ide_action.action_args)
                    else:
                        raise ValueError("save_file 动作参数错误")
                    return IDEObs(obs="保存成功").model_dump(), 100, True, True, {}
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "close_file":
                try:
                    if isinstance(ide_action.action_args, dict):
                        self.close_file(**ide_action.action_args)
                    elif isinstance(ide_action.action_args, str):
                        self.close_file(uri=ide_action.action_args)
                    else:
                        raise ValueError("close_file 动作参数错误")
                    return IDEObs(obs="关闭成功").model_dump(), 100, True, True, {}
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "read_file":
                try:
                    if isinstance(ide_action.action_args, dict):
                        # 构建调用参数
                        if ide_action.action_args.get("code_range"):
                            ide_action.action_args["code_range"] = Range.model_validate(
                                ide_action.action_args["code_range"],
                            )
                        text = self.read_file(**ide_action.action_args)
                    elif isinstance(ide_action.action_args, str):
                        text = self.read_file(uri=ide_action.action_args)
                    else:
                        raise ValueError("read_file 动作参数错误")
                    return IDEObs(obs=text).model_dump(), 100, True, True, {}
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "get_file_symbols":
                try:
                    if isinstance(ide_action.action_args, dict):
                        symbols = self.get_file_symbols(**ide_action.action_args)
                    else:
                        raise ValueError("get_file_symbols 动作参数错误")
                    return IDEObs(obs=symbols).model_dump(), 100, True, True, {}
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "find_in_file":
                try:
                    if isinstance(ide_action.action_args, dict):
                        if ide_action.action_args.get("search_scope"):
                            if isinstance(ide_action.action_args["search_scope"], dict):
                                ide_action.action_args["search_scope"] = Range.model_validate(
                                    ide_action.action_args["search_scope"],
                                )
                            elif isinstance(ide_action.action_args["search_scope"], list):
                                ide_action.action_args["search_scope"] = [
                                    Range.model_validate(r) for r in ide_action.action_args["search_scope"]
                                ]
                        search_res = self.find_in_path(**ide_action.action_args)
                    else:
                        raise ValueError("find_in_file 动作参数错误")
                    return (
                        IDEObs(obs="\n".join([repr(r) for r in search_res])).model_dump(),
                        100,
                        True,
                        True,
                        {},
                    )
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "replace_in_file":
                try:
                    if isinstance(ide_action.action_args, dict):
                        if ide_action.action_args.get("search_scope"):
                            if isinstance(ide_action.action_args["search_scope"], dict):
                                ide_action.action_args["search_scope"] = Range.model_validate(
                                    ide_action.action_args["search_scope"],
                                )
                            elif isinstance(ide_action.action_args["search_scope"], list):
                                ide_action.action_args["search_scope"] = [
                                    Range.model_validate(r) for r in ide_action.action_args["search_scope"]
                                ]
                        undo_edits, diagnostics = self.replace_in_file(**ide_action.action_args)
                    else:
                        raise ValueError("replace_in_file 动作参数错误")

                    # 构建观察结果 / Build observation result
                    obs_text = "完成替换"
                    if diagnostics:
                        obs_text += self._format_diagnostics(diagnostics)

                    return (
                        IDEObs(
                            obs=obs_text,
                            original_result={"undo_edits": undo_edits, "diagnostics": diagnostics},
                        ).model_dump(),
                        100,
                        True,
                        True,
                        {},
                    )
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "create_file":
                try:
                    if isinstance(ide_action.action_args, dict):
                        create_model, diagnostics = self.create_file(**ide_action.action_args)
                    elif isinstance(ide_action.action_args, str):
                        create_model, diagnostics = self.create_file(uri=ide_action.action_args)
                    else:
                        raise ValueError("create_file 动作参数错误")
                    if create_model:
                        obs_text = (
                            "文件创建成功。\n当前文件内容如下(IDE会自动初始化部分内容):\n"
                            f"{self.read_file(uri=str(create_model.uri), with_line_num=True)}\n"
                            if create_model.get_value()
                            else "文件创建成功"
                        )
                        # 添加诊断信息 / Add diagnostics
                        if diagnostics:
                            obs_text += self._format_diagnostics(diagnostics)

                        return (
                            IDEObs(obs=obs_text).model_dump(),
                            100,
                            True,
                            True,
                            {},
                        )
                    else:
                        return IDEObs(obs="文件已存在").model_dump(), 0, True, False, {}
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "insert_cursor":
                try:
                    if isinstance(ide_action.action_args, dict):
                        insert_res = self.insert_cursor(**ide_action.action_args)
                    else:
                        raise ValueError("insert_cursor 动作参数错误")
                    return IDEObs(obs=insert_res).model_dump(), 100, True, True, {}
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "delete_cursor":
                try:
                    if isinstance(ide_action.action_args, dict):
                        delete_res = self.delete_cursor(**ide_action.action_args)
                    else:
                        raise ValueError("delete_cursor 动作参数错误")
                    return IDEObs(obs=delete_res).model_dump(), 100, True, True, {}
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case "clear_cursors":
                try:
                    if isinstance(ide_action.action_args, dict):
                        clear_res = self.clear_cursors(**ide_action.action_args)
                    else:
                        raise ValueError("clear_cursors 动作参数错误")
                    return IDEObs(obs=clear_res).model_dump(), 100, True, True, {}
                except Exception as e:
                    return IDEObs(obs=str(e)).model_dump(), 0, True, False, {}
            case (
                "get_definition_and_implementation"
                | "hover"
                | "find_in_workspace"
                | "replace_in_workspace"
                | "create_files"
                | "delete_files"
                | "rename_file"
                | "delete_file"
            ):
                raise NotImplementedError(f"Action: {ide_action.action_name} 尚未实现")
            case _:
                raise ValueError(f"不支持的动作 {ide_action.action_name}")  # pragma: no cover

    def render(self, *, verbose: bool = False) -> str:  # type: ignore
        """
        渲染当前工作区状态，主要提取active_models相关信息，最后一个active_models取其view全部内容，之前的active_models取其symbols信息
        Render current workspace state, extract active_models info, show full view for last model and symbols for others

        Args:
            verbose (bool): 是否使用详细模式。True时返回包含Python包/模块描述的丰富信息，False时返回简化版本
                           | Whether to use verbose mode. True returns rich info with Python package/module descriptions,
                           False returns simplified version

        Returns:
            str: 以字符串的形式来返回渲染结果 | Render result as string
        """
        self._assert_not_closed()

        # 1. 渲染最小化展开的目录树 | Render minimally expanded directory tree
        # 根据verbose参数选择使用简化版本或丰富版本 | Choose simplified or rich version based on verbose parameter
        if verbose:
            from ide4ai.python_ide.utils import get_minimal_expanded_tree_with_desc as get_minimal_tree
            from ide4ai.python_ide.utils import list_directory_tree_with_desc as list_dir_tree
        else:
            from ide4ai.utils import get_minimal_expanded_tree as get_minimal_tree  # type: ignore[assignment]
            from ide4ai.utils import list_directory_tree as list_dir_tree  # type: ignore[assignment]

        if self.active_models:
            # 获取最后一个active_model的文件路径 | Get the last active_model's file path
            last_model = self.active_models[-1]
            target_file = str(last_model.uri).replace("file://", "")
            dir_info = get_minimal_tree(self.root_dir, target_file, indent="- ")
        else:
            # 如果没有active_models，使用普通的目录树 | Use normal directory tree if no active_models
            dir_info = list_dir_tree(self.root_dir, include_dirs=self.expand_folders, recursive=True, indent="- ")

        view = f"当前工作区: {self.project_name}\n\n项目目录结构:\n{dir_info}\n"

        # 2. 添加项目快捷命令信息 | Add project shortcut commands info
        if self.shortcut_commands:
            view += "\n项目快捷命令 | Project Shortcut Commands:\n"
            for cmd_prefix, cmd_list in self.shortcut_commands.items():
                view += f"  {cmd_prefix} 命令:\n"
                for cmd in cmd_list:
                    view += f"    - {cmd_prefix} {cmd}\n"
            view += "\n"
        else:
            # 如果没有shortcut_commands，尝试实时检测Makefile | Try to detect Makefile in real-time if no shortcut_commands
            from ide4ai.utils import detect_makefile_commands

            detected_commands = detect_makefile_commands(self.root_dir)
            if detected_commands:
                view += "\n项目快捷命令 | Project Shortcut Commands:\n"
                for cmd_prefix, cmd_list in detected_commands.items():
                    view += f"  {cmd_prefix} 命令:\n"
                    for cmd in cmd_list:
                        view += f"    - {cmd_prefix} {cmd}\n"
                view += "\n"

        # 3. 渲染active_models信息 | Render active_models info
        active_models_count = len(self.active_models)
        if active_models_count > 1:
            view += "\n以下是最近使用的文件其结构信息与关键Symbols信息。每个结构跟随一个Range范围，可以使用这个Range+URI查询代码详情:\n"
            for active_view in self.active_models[:-1]:
                uri = active_view.uri
                view += f"文件URI: {uri}\n"
                mid = self.get_lsp_msg_id()
                lsp_res = self.send_lsp_msg(
                    "textDocument/documentSymbol",
                    {"textDocument": {"uri": str(uri)}},
                    message_id=mid,
                )
                if lsp_res:
                    res_model = LSPResponseMessage.model_validate(json.loads(lsp_res))
                    if res_model.error:
                        view += f"获取Symbols信息失败: {res_model.error}\n"  # pragma: no cover
                        continue  # pragma: no cover
                    symbols = res_model.result
                    view += render_symbols(symbols, DEFAULT_SYMBOL_VALUE_SET)  # type: ignore
                    view += "\n"
                else:
                    view += "无法获取Symbols信息\n"  # pragma: no cover
        if active_models_count > 0:
            view += f"当前打开的文件内容如下：\n{self.active_models[-1].get_view()}\n"
        return view

    def open_file(self, *, uri: str) -> TextModel:
        """
        Open a file in the workspace.
        Initial a model instance, add it to self.models and active it

        Args:
            uri (str): The path to the file to be opened.

        Returns:
            TextModel: The model instance representing the opened file.
        """
        self._assert_not_closed()
        if tm := next(filter(lambda model: model.uri == AnyUrl(uri), self.models), None):
            self.active_model(tm.m_id)  # pragma: no cover
            return tm  # pragma: no cover
        text_model = TextModel(language_id=LanguageId.python, uri=AnyUrl(uri))
        self.models.append(text_model)
        self.active_model(text_model.m_id)
        self.send_lsp_msg(
            "textDocument/didOpen",
            {
                "textDocument": {
                    "uri": uri,
                    "languageId": LanguageId.python.value,
                    "version": text_model.get_version_id(),
                    "text": text_model.get_value(),
                },
            },
        )
        return text_model

    def apply_edit(
        self,
        *,
        uri: str,
        edits: Sequence[SingleEditOperation | dict],
        compute_undo_edits: bool = False,
    ) -> tuple[list[TextEdit] | None, DocumentDiagnosticReport | None]:
        """
        Apply edits to a file in the workspace.

        Notes:
            注意这个函数不支持 "-1" 形式的Range调用，即不能用负数表示倒数。如果想实现需要在外侧完成转换后再传入。
            编辑后会自动拉取诊断信息 / Diagnostics will be automatically pulled after editing

        Args:
            uri (str): The URI of the file to which the edits should be applied.
            edits (list[SingleEditOperation | dict]): The edits to be applied to the file.
            compute_undo_edits (bool): Whether to compute the undo edits. This parameter is optional and defaults to
                False.

        Returns:
            tuple[Optional[list[TextEdit]], Optional[DocumentDiagnosticReport]]:
                - The reverse edits that can be applied to undo the changes / 可用于撤销更改的反向编辑
                - Diagnostics result after editing / 编辑后的诊断结果
        """
        self._assert_not_closed()
        text_model = next(filter(lambda model: model.uri == AnyUrl(uri), self.models), None)
        if not text_model:
            text_model = self.open_file(uri=uri)  # pragma: no cover
        try:
            model_edits = [
                SingleEditOperation.model_validate(edit) if isinstance(edit, dict) else edit for edit in edits
            ]
        except ValidationError as e:
            err_info = (
                f"编辑操作参数错误，具体报错如下:\n{e}\n这类错误经常由Range范围引起，在当前工作区内Range与Position均是1-based。"
                f"不要使用0基索引"
            )
            raise IDEExecutionError(message=err_info, detail_for_llm=err_info) from e
        res = text_model.apply_edits(model_edits, compute_undo_edits)

        self.send_lsp_msg(
            "textDocument/didChange",
            {
                "textDocument": {"uri": uri, "version": text_model.get_version_id()},
                "contentChanges": [
                    {"range": edit.range.to_lsp_range().model_dump(), "text": edit.text} for edit in model_edits
                ],
            },
        )

        # 编辑后主动拉取诊断信息 / Pull diagnostics after editing
        diagnostics = self.pull_diagnostics(uri=uri, timeout=self._diagnostics_timeout)

        return res, cast(
            DocumentDiagnosticReport | None,
            diagnostics,
        )  # 在指定uri的情况下，返回的diagnostics为DocumentDiagnostics

    def rename_file(
        self,
        *,
        old_uri: str,
        new_uri: str,
        overwrite: bool | None = None,
        ignore_if_exists: bool | None = None,
    ) -> bool:
        """
        # TODO 需要与LSP进行信息互通查询到相应的引用关系后，将引用关系变更后，再进行文件重命名。这个过程涉及到LSP互通与异常回滚。目前未实现
        Rename a file.

        Args:
            old_uri:
            new_uri:
            overwrite:
            ignore_if_exists:

        Returns:
            bool: True if the file is successfully renamed, False otherwise.
        """
        raise NotImplementedError(
            "rename_file 需要与LSP进行信息互通查询到相应的引用关系后，将引用关系变更后，再进行文件重命名。这个过程涉及到LSP互通与异常回滚。"
            "目前暂未实现。你可以提示用户使用PyCharm等工具手动重命名文件。",
        )

    def delete_file(
        self,
        *,
        uri: str,
        recursive: bool | None = None,
        ignore_if_not_exists: bool | None = None,
    ) -> bool:
        """
        Deletes a file from the specified URI.

        Args:
            uri: The URI of the file to be deleted.
            recursive: Optional. If set to True, deletes the file recursively along with any directories. If set to
                False, only deletes the file.
            ignore_if_not_exists: Optional. If set to True, no error will be raised if the file does not exist. If set
                to False, an error will be raised if the file does not exist.

        Returns:
            bool: True if the file is successfully deleted, False otherwise.

        Raises:
            NotImplementedError: This method is not implemented yet. It requires communication with LSP to update the
                reference relationships before renaming the file. This process involves LSP communication and exception
                rollback. Currently, it is not implemented. You can suggest the user to manually delete the file using
                tools like PyCharm.

        Example:
            delete_file(uri='/path/to/file.txt')
        """
        # TODO 需要与LSP进行信息互通查询到相应的引用关系后，将引用关系变更后，再进行文件删除。这个过程涉及到LSP互通与异常回滚。目前未实现
        raise NotImplementedError(
            "delete_file 需要与LSP进行信息互通查询到相应的引用关系后，将引用关系变更后，再进行文件删除。这个过程涉及到LSP互通与异常回滚。"
            "目前暂未实现。你可以提示用户使用PyCharm等工具手动删除文件。",
        )

    def create_file(
        self,
        *,
        uri: str,
        init_content: str | None = None,
        overwrite: bool | None = None,
        ignore_if_exists: bool | None = None,
    ) -> tuple[TextModel | None, DocumentDiagnosticReport | None]:
        """
        Create a file at the specified URI.

        Args:
            uri (str): The path where the file will be created.
            init_content (str, optional): The initial content of the file. Defaults to None.
            overwrite (bool, optional): If True, overwrite the file if it exists. Defaults to None.
            ignore_if_exists (bool, optional): If True, do nothing if the file already exists. Defaults to None.

        Returns:
            tuple[Optional[TextModel], Optional[DocumentDiagnosticReport]]:
                - The model instance representing the created file / 创建的文件模型实例
                - Diagnostics result after creation / 创建后的诊断结果
        """
        self._assert_not_closed()
        if not uri.startswith("file://"):
            uri = f"file://{uri}"  # pragma: no cover
        file_path = uri[7:]

        # Check if the file already exists
        if os.path.exists(file_path):
            if ignore_if_exists:
                return None, None  # Do nothing as the file exists and we should ignore this situation
            if not overwrite:
                raise FileExistsError(f"The file at {file_path} already exists and overwrite is not set to True.")
            # Overwrite is True, delete the file before creating a new one
            os.remove(file_path)
        # Pyright 目前好像不支持 workspace/willCreateFiles 方法
        # msg_id = self.get_lsp_msg_id()
        # lsp_res_will_create = self.send_lsp_msg("workspace/willCreateFiles", {"files": [{"uri": uri}]}, msg_id)
        # if not lsp_res_will_create:
        #     raise ValueError(f"无法创建文件: {uri}， LSP校验未通过")
        # lsp_res = LSPResponseMessage.model_validate(json.loads(lsp_res_will_create))
        # if lsp_res.error:
        #     raise ValueError(f"无法创建文件: {uri}， LSP校验未通过: {lsp_res.error}")
        # TODO LSP会响应 workspace/willCreateFiles Request，返回的结构中会包括一个workspaceEdit。 \
        #  完成apply_workspace方法的封装后，需要在此调用并响应
        # Create the file
        try:
            # Using 'x' mode to create file will raise an error if the file already exists
            with open(file_path, "x") as file:
                for file_type, header_generator in self.header_generators.items():
                    if file_path.endswith(file_type):
                        header = header_generator(self, file_path)
                        file.write(header)
                        break
            tm = TextModel(language_id=LanguageId.python, uri=AnyUrl(uri))

            # 在文件创建后追加初始化内容（如果存在）/ Append initial content after file creation (if exists)
            if init_content:
                tm.apply_edits(
                    [
                        SingleEditOperation(
                            range=Range(
                                start_position=Position(
                                    tm.get_line_count(),
                                    tm.get_line_length(tm.get_line_count()) + 1,
                                ),
                                end_position=Position(tm.get_line_count(), tm.get_line_length(tm.get_line_count()) + 1),
                            ),
                            text=("\n" + init_content) if tm.get_line_count() > 1 else init_content,
                        ),
                    ],
                )

            self.models.append(tm)
            self.active_model(tm.m_id)

            # 通知LSP文件已创建 / Notify LSP that file has been created
            self.send_lsp_msg("workspace/didCreateFiles", {"files": [{"uri": uri}]})

            # 通知LSP打开文件，发送完整内容（包含header和init_content）/ Notify LSP to open file with complete content
            self.send_lsp_msg(
                "textDocument/didOpen",
                {
                    "textDocument": {
                        "uri": uri,
                        "languageId": LanguageId.python.value,
                        "version": tm.get_version_id(),
                        "text": tm.get_value(),
                    },
                },
            )

            # 创建文件后主动拉取诊断信息 / Pull diagnostics after file creation
            diagnostics = self.pull_diagnostics(uri=uri, timeout=5.0)

            return tm, cast(DocumentDiagnosticReport | None, diagnostics)
        except FileExistsError:
            # If overwrite was True, we already deleted the file, so this should not happen
            return None, None  # pragma: no cover
        except Exception as e:
            # Handle other possible exceptions, such as permission errors
            raise OSError(f"Failed to create file at {uri}: {str(e)}") from e

    def find_in_path(
        self,
        *,
        uri: str,
        query: str,
        search_scope: Range | list[Range] | None = None,
        is_regex: bool = False,
        match_case: bool = False,
        word_separator: str | None = None,
        capture_matches: bool = True,
        limit_result_count: int | None = None,
    ) -> list[SearchResult]:
        """
        在工作区中的文件或文件夹内查找查询字符串 / Find a query in a file or folder in the workspace.

        Args:
            uri (str): 要搜索的文件或文件夹的 URI。如果是文件夹，将递归搜索其中的所有文件 /
                      The URI of the file or folder to search in. If it's a folder, will recursively search all files within.
            query (str): 要搜索的查询字符串 / The query to search for.
            search_scope: 可选。指定搜索应在其中进行的范围或范围列表。仅当 uri 是文件时有效。如果未提供，
                则在整个文件范围内进行搜索 / Optional. The range or list of ranges where the search should be performed.
                Only valid when uri is a file. If not provided, the search will be performed in the full file range.
            is_regex: 可选。指定是否应将搜索字符串视为正则表达式。默认为 False /
                     Optional. Specifies whether the search string should be treated as a regular expression. Default is False.
            match_case: 可选。指定搜索是否应区分大小写。默认为 False /
                       Optional. Specifies whether the search should be case-sensitive. Default is False.
            word_separator: 可选。用于定义搜索中单词边界的分隔符。如果未提供，则所有字符都视为单词的一部分 /
                          Optional. The separator used to define word boundaries in the search. If not provided,
                          all characters are considered as part of a word.
            capture_matches: 可选。指定是否应在搜索结果中捕获匹配的文本内容。默认为 True /
                           Optional. Specifies whether the matched text should be captured in the search results. Default is True.
            limit_result_count: 可选。返回的搜索结果的最大数量。如果未提供，将返回所有匹配项 /
                              Optional. The maximum number of search results to return. If not provided, all matches will be returned.

        Returns:
            表示匹配结果的 SearchResult 对象列表。每个结果包含匹配的范围和文本（如果 capture_matches 为 True）/
            A list of SearchResult objects representing the matched results. Each result contains the matched range
            and text (if capture_matches is True).

        Raises:
            ValueError: 如果提供了无效的 URI 或搜索范围 / If an invalid URI or search scope is provided.

        Examples:
            # 在单个文件中搜索 / Search in a single file
            results = workspace.find_in_path(uri="file:///path/to/file.py", query="def")

            # 在文件夹中递归搜索 / Recursively search in a folder
            results = workspace.find_in_path(uri="file:///path/to/folder", query="TODO", match_case=True)

            # 使用正则表达式搜索 / Search with regex
            results = workspace.find_in_path(uri="file:///path/to/file.py", query=r"\\bclass\\s+\\w+", is_regex=True)
        """
        # 验证 URI 格式 / Validate URI format
        if not uri.startswith("file://"):
            raise ValueError(f"URI 必须以 'file://' 开头 / URI must start with 'file://': {uri}")

        # 提取文件路径 / Extract file path
        from pathlib import Path

        file_path = Path(uri[7:])

        # 验证路径存在 / Validate path exists
        if not file_path.exists():
            raise ValueError(f"路径不存在 / Path does not exist: {file_path}")

        # 如果是文件，直接搜索 / If it's a file, search directly
        if file_path.is_file():
            text_model = self.get_model(uri)
            if not text_model:
                text_model = TextModel(language_id=LanguageId.python, uri=AnyUrl(uri))
            return text_model.find_matches(
                query,
                search_scope,
                is_regex,
                match_case,
                word_separator,
                capture_matches,
                limit_result_count,
            )

        # 如果是文件夹，递归搜索所有文件 / If it's a folder, recursively search all files
        if file_path.is_dir():
            # 当搜索文件夹时，search_scope 参数无效 / search_scope is not valid when searching a folder
            if search_scope is not None:
                raise ValueError(
                    "搜索文件夹时不支持 search_scope 参数 / search_scope is not supported when searching a folder",
                )

            all_results: list[SearchResult] = []
            result_count = 0

            # 递归遍历文件夹中的所有文件 / Recursively iterate all files in the folder
            for file in file_path.rglob("*"):
                if file.is_file():
                    try:
                        file_uri = f"file://{file.absolute()}"

                        # 先尝试从已打开的模型中获取 / Try to get from opened models first
                        text_model = self.get_model(file_uri)

                        if not text_model:
                            # 如果文件未打开，创建临时模型进行搜索 / If file is not opened, create temporary model for search
                            try:
                                # 检查文件是否可读 / Check if file is readable
                                if not file.exists() or not file.is_file():
                                    continue

                                # 创建临时文本模型 / Create temporary text model
                                text_model = TextModel(
                                    language_id=LanguageId.python,
                                    uri=AnyUrl(file_uri),
                                    auto_save_during_dispose=False,  # 临时模型不需要自动保存 / Temporary model doesn't need auto-save
                                )
                            except Exception as e:
                                # 跳过无法读取的文件 / Skip files that cannot be read
                                from loguru import logger

                                logger.debug(f"跳过文件 / Skip file {file}: {e}")
                                continue

                        # 计算当前文件的结果限制 / Calculate result limit for current file
                        current_limit = None
                        if limit_result_count is not None:
                            remaining = limit_result_count - result_count
                            if remaining <= 0:
                                break
                            current_limit = remaining

                        # 在当前文件中搜索 / Search in current file
                        file_results = text_model.find_matches(
                            query,
                            None,  # 文件夹搜索不使用 search_scope / No search_scope for folder search
                            is_regex,
                            match_case,
                            word_separator,
                            capture_matches,
                            current_limit,
                        )

                        all_results.extend(file_results)
                        result_count += len(file_results)

                        # 如果达到结果数量限制，停止搜索 / Stop if result limit is reached
                        if limit_result_count is not None and result_count >= limit_result_count:
                            break

                    except Exception as e:
                        # 跳过处理失败的文件 / Skip files that fail to process
                        from loguru import logger

                        logger.debug(f"处理文件失败 / Failed to process file {file}: {e}")
                        continue

            return all_results

        # 如果既不是文件也不是文件夹 / If it's neither a file nor a folder
        raise ValueError(f"URI 必须指向文件或文件夹 / URI must point to a file or folder: {uri}")

    def grep_files(
        self,
        *,
        pattern: str,
        path: str | None = None,
        glob: str | None = None,
        output_mode: str = "files_with_matches",
        context_before: int | None = None,
        context_after: int | None = None,
        context: int | None = None,
        line_number: bool | None = None,
        case_insensitive: bool | None = None,
        file_type: str | None = None,
        head_limit: int | None = None,
        multiline: bool = False,
    ) -> dict[str, Any]:
        """
        使用 ripgrep 搜索文件内容 / Search file contents using ripgrep

        Args:
            pattern: 正则表达式模式 / Regular expression pattern
            path: 搜索路径，默认为工作区根目录 / Search path, defaults to workspace root
            glob: 文件过滤 Glob 模式 / Glob pattern to filter files
            output_mode: 输出模式 / Output mode: "content", "files_with_matches", or "count"
            context_before: 显示匹配前的行数 / Lines before match
            context_after: 显示匹配后的行数 / Lines after match
            context: 显示匹配前后的行数 / Lines before and after match
            line_number: 显示行号 / Show line numbers
            case_insensitive: 忽略大小写 / Case insensitive search
            file_type: 文件类型 / File type to search
            head_limit: 限制输出行数 / Limit output to first N lines
            multiline: 启用多行模式 / Enable multiline mode

        Returns:
            dict: 搜索结果 / Search results

        Examples:
            # 搜索所有包含 "TODO" 的文件
            workspace.grep_files(pattern="TODO")

            # 搜索 Python 文件中的类定义
            workspace.grep_files(pattern="class\\s+\\w+", file_type="py", output_mode="content")

            # 多行搜索
            workspace.grep_files(pattern="def.*\\n.*return", multiline=True)
        """
        self._assert_not_closed()

        # 确定搜索路径 / Determine search path
        search_path = path if path else self.root_dir

        # 如果是相对路径，转换为相对于工作区根目录的绝对路径 / If relative path, convert to absolute
        if not os.path.isabs(search_path):
            search_path = os.path.join(self.root_dir, search_path)

        # 验证路径存在 / Validate path exists
        if not os.path.exists(search_path):
            raise ValueError(f"搜索路径不存在 / Search path does not exist: {search_path}")

        # 确保搜索路径在工作区内 / Ensure search path is within workspace
        from ide4ai.utils import is_subdirectory

        if not is_subdirectory(search_path, self.root_dir):
            raise ValueError(f"搜索路径必须在工作区根目录内 / Search path must be within workspace root: {search_path}")

        # 构建 ripgrep 命令 / Build ripgrep command
        cmd = ["rg"]

        # 添加输出模式参数 / Add output mode parameters
        if output_mode == "files_with_matches":
            cmd.append("--files-with-matches")
        elif output_mode == "count":
            cmd.append("--count")
        # content 模式不需要额外参数 / content mode needs no extra parameters

        # 添加上下文参数（仅在 content 模式下有效）/ Add context parameters (only valid in content mode)
        if output_mode == "content":
            if context is not None:
                cmd.extend(["-C", str(context)])
            else:
                if context_before is not None:
                    cmd.extend(["-B", str(context_before)])
                if context_after is not None:
                    cmd.extend(["-A", str(context_after)])

            # 添加行号参数 / Add line number parameter
            if line_number:
                cmd.append("-n")

        # 添加大小写敏感参数 / Add case sensitivity parameter
        if case_insensitive:
            cmd.append("-i")

        # 添加文件类型参数 / Add file type parameter
        if file_type:
            cmd.extend(["--type", file_type])

        # 添加 glob 参数 / Add glob parameter
        if glob:
            cmd.extend(["--glob", glob])

        # 添加多行模式参数 / Add multiline mode parameters
        if multiline:
            cmd.extend(["-U", "--multiline-dotall"])

        # 添加模式和路径 / Add pattern and path
        cmd.append(pattern)
        cmd.append(search_path)

        try:
            # 执行 ripgrep 命令 / Execute ripgrep command
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=30,  # 30秒超时 / 30 seconds timeout
            )

            # ripgrep 退出码：0=找到匹配，1=未找到匹配，2=错误
            # ripgrep exit codes: 0=found matches, 1=no matches, 2=error
            if result.returncode == 2:
                raise RuntimeError(f"ripgrep 执行错误 / ripgrep execution error: {result.stderr}")

            output = result.stdout

            # 应用 head_limit / Apply head_limit
            if head_limit is not None and output:
                lines = output.splitlines()
                output = "\n".join(lines[:head_limit])

            return {
                "success": True,
                "output": output,
                "matched": result.returncode == 0,
                "metadata": {
                    "pattern": pattern,
                    "path": search_path,
                    "output_mode": output_mode,
                    "exit_code": result.returncode,
                },
            }

        except subprocess.TimeoutExpired as e:
            raise RuntimeError("ripgrep 执行超时 / ripgrep execution timeout") from e
        except FileNotFoundError as e:
            raise RuntimeError(
                "ripgrep 未安装。请安装 ripgrep: https://github.com/BurntSushi/ripgrep#installation"
            ) from e
        except Exception as e:
            raise RuntimeError(f"执行 ripgrep 时发生错误 / Error executing ripgrep: {e}") from e

    def apply_workspace_edit(self, *, workspace_edit: LSPWorkspaceEdit) -> Any:
        # TODO 需要实现 apply_workspace_edit 方法
        raise NotImplementedError("apply_workspace_edit 尚未实现")
