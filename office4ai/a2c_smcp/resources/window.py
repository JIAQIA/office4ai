"""window://office4ai Resource — Office Workspace 根索引资源."""

from __future__ import annotations

from urllib.parse import urlencode

from office4ai.a2c_smcp.resources.base import BaseResource, parse_window_uri_params
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import connection_manager


class WindowResource(BaseResource):
    """
    Office Workspace 根索引资源

    通过 ``window://office4ai`` 向 AI Agent 展示子资源索引总览，
    按文档类型统计已连接文档数。

    渲染示例（有文档连接时）::

        # Office 工作区

        ## 子资源
        - window://office4ai/word — Word 文档 (2 个已连接)
        - window://office4ai/ppt — PPT 文档 (1 个已连接)

    渲染示例（无文档连接时）::

        # Office 工作区

        ## 子资源
        暂无文档连接，等待 Office Add-In 接入。
    """

    def __init__(self, workspace: OfficeWorkspace, priority: int = 0, fullscreen: bool = False) -> None:
        if not isinstance(priority, int) or not (0 <= priority <= 100):
            raise ValueError(f"priority must be int in [0, 100], got: {priority}")
        self.workspace = workspace
        self._priority = priority
        self._fullscreen = fullscreen

    # ── BaseResource implementation ──

    @property
    def uri(self) -> str:
        query = urlencode(
            {
                "priority": str(self._priority),
                "fullscreen": "true" if self._fullscreen else "false",
            }
        )
        return f"window://office4ai?{query}"

    @property
    def base_uri(self) -> str:
        return "window://office4ai"

    @property
    def name(self) -> str:
        return "Office 工作区"

    @property
    def description(self) -> str:
        return "Office 工作区根索引，展示子资源（Word/PPT）连接状态总览。"

    @property
    def mime_type(self) -> str:
        return "text/plain"

    async def read(self) -> str:
        return self._render()

    def update_from_uri(self, uri: str) -> None:
        self._priority, self._fullscreen = parse_window_uri_params(
            uri, self._priority, self._fullscreen, log_prefix="Window resource"
        )

    # ── Rendering ──

    def _render(self) -> str:
        clients = connection_manager.get_all_clients()

        # 按 namespace 统计文档数（按 document_uri 去重）
        word_docs: set[str] = set()
        ppt_docs: set[str] = set()
        for c in clients:
            if c.namespace == "/word":
                word_docs.add(c.document_uri)
            elif c.namespace == "/ppt":
                ppt_docs.add(c.document_uri)

        lines: list[str] = ["# Office 工作区", ""]
        lines.append("## 子资源")

        if word_docs or ppt_docs:
            lines.append(f"- window://office4ai/word — Word 文档 ({len(word_docs)} 个已连接)")
            lines.append(f"- window://office4ai/ppt — PPT 文档 ({len(ppt_docs)} 个已连接)")
        else:
            lines.append("暂无文档连接，等待 Office Add-In 接入。")

        return "\n".join(lines)
