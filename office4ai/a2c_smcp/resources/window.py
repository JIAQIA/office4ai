"""window://office4ai Resource — Office Workspace window for A2C Desktop."""

from __future__ import annotations

import time
from urllib.parse import parse_qs, urlencode, urlparse

from loguru import logger

from office4ai.a2c_smcp.resources.base import BaseResource
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import connection_manager


class WindowResource(BaseResource):
    """
    Office Workspace 窗口资源

    通过 A2C-SMCP 定义的 ``window://`` 协议，向 AI Agent 暴露已连接文档、
    活跃文档及缓存的可见内容/文档结构。

    渲染示例（有文档连接时）::

        Office 工作区
        =============
        已连接文档 (2):
          [word] file:///Users/jqq/Documents/report.docx
          [word] file:///Users/jqq/Documents/draft.docx

        活跃文档: file:///Users/jqq/Documents/report.docx
        最近工具: word_get_visible_content (2秒前)

        --- 可见内容 ---
        <最近一次 word_get_visible_content 返回的缓存内容>

        --- 文档结构 ---
        <最近一次 word_get_document_structure 返回的缓存结构>

    渲染示例（无文档连接时）::

        Office 工作区
        =============
        已连接文档 (0):
          暂无文档连接，等待 Office Add-In 接入。
    """

    def __init__(self, workspace: OfficeWorkspace, priority: int = 0, fullscreen: bool = True) -> None:
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
        return "Office 工作区窗口，展示已连接文档、活跃文档及缓存的可见内容/文档结构。"

    @property
    def mime_type(self) -> str:
        return "text/plain"

    async def read(self) -> str:
        return self._render()

    def update_from_uri(self, uri: str) -> None:
        parsed = urlparse(uri)
        params = parse_qs(parsed.query)

        if "priority" in params:
            try:
                new_priority = int(params["priority"][0])
                if 0 <= new_priority <= 100:
                    if new_priority != self._priority:
                        logger.debug(f"Window resource priority: {self._priority} -> {new_priority}")
                        self._priority = new_priority
                else:
                    logger.warning(f"Invalid priority value in URI: {new_priority}, must be in [0, 100]")
            except (ValueError, IndexError) as e:
                logger.warning(f"Failed to parse priority from URI: {e}")

        if "fullscreen" in params:
            try:
                fs_str = params["fullscreen"][0].lower()
                new_fs: bool | None = None
                if fs_str in {"true", "1", "yes", "on"}:
                    new_fs = True
                elif fs_str in {"false", "0", "no", "off"}:
                    new_fs = False
                else:
                    logger.warning(f"Invalid fullscreen value in URI: {fs_str}, ignoring")

                if new_fs is not None and new_fs != self._fullscreen:
                    logger.debug(f"Window resource fullscreen: {self._fullscreen} -> {new_fs}")
                    self._fullscreen = new_fs
            except (IndexError, AttributeError) as e:
                logger.warning(f"Failed to parse fullscreen from URI: {e}")

    # ── Rendering ──

    def _render(self) -> str:
        clients = connection_manager.get_all_clients()
        last = self.workspace.get_last_activity()

        lines: list[str] = [
            "Office 工作区",
            "=============",
        ]

        # 已连接文档区域 — 按 document_uri 去重，保留命名空间信息
        doc_map: dict[str, str] = {}  # document_uri → namespace (strip leading /)
        for c in clients:
            ns_label = c.namespace.lstrip("/") if c.namespace else "unknown"
            doc_map[c.document_uri] = ns_label

        lines.append(f"已连接文档 ({len(doc_map)}):")
        if doc_map:
            for doc_uri, ns_label in doc_map.items():
                lines.append(f"  [{ns_label}] {doc_uri}")
        else:
            lines.append("  暂无文档连接，等待 Office Add-In 接入。")

        # 活跃文档区域
        if last is not None:
            elapsed = time.time() - last.timestamp
            elapsed_str = f"{elapsed:.0f}秒前" if elapsed < 60 else f"{elapsed / 60:.1f}分钟前"
            lines.append("")
            lines.append(f"活跃文档: {last.document_uri}")
            lines.append(f"最近工具: {last.tool_name} ({elapsed_str})")

            # 缓存的可见内容
            cached_content = self.workspace.get_cached_content(last.document_uri)
            if cached_content:
                lines.append("")
                lines.append("--- 可见内容 ---")
                lines.append(cached_content)

            # 缓存的文档结构
            cached_structure = self.workspace.get_cached_structure(last.document_uri)
            if cached_structure:
                lines.append("")
                lines.append("--- 文档结构 ---")
                lines.append(cached_structure)

        return "\n".join(lines)
