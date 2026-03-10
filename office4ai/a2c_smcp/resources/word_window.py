"""window://office4ai/word Resource — Word 文档聚合窗口资源."""

from __future__ import annotations

import asyncio
from typing import Any
from urllib.parse import urlencode

from loguru import logger

from office4ai.a2c_smcp.resources.base import BaseResource, parse_window_uri_params
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import connection_manager


class WordWindowResource(BaseResource):
    """
    Word 文档聚合窗口资源

    通过 ``window://office4ai/word`` 向 AI Agent 暴露 Word 文档的实时状态，
    包括已连接文档列表、激活文档的元数据和可见内容。

    每次 ``read()`` 通过 Socket.IO 拉取最新数据，3 秒超时后降级渲染。
    """

    FETCH_TIMEOUT = 3  # 秒

    def __init__(self, workspace: OfficeWorkspace, priority: int = 50, fullscreen: bool = True) -> None:
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
        return f"window://office4ai/word?{query}"

    @property
    def base_uri(self) -> str:
        return "window://office4ai/word"

    @property
    def name(self) -> str:
        return "Word 工作区"

    @property
    def description(self) -> str:
        return "Word 文档聚合窗口，展示已连接 Word 文档列表、激活文档的元数据和可见内容。"

    @property
    def mime_type(self) -> str:
        return "text/plain"

    async def read(self) -> str:
        clients = connection_manager.get_all_clients()
        last = self.workspace.get_last_activity()

        # 过滤 /word namespace 文档，按 document_uri 去重
        word_docs: set[str] = set()
        for c in clients:
            if c.namespace == "/word":
                word_docs.add(c.document_uri)

        # 确定激活文档（last_activity 必须也在 /word namespace）
        active_uri: str | None = None
        if last is not None and last.document_uri in word_docs:
            active_uri = last.document_uri

        lines: list[str] = ["# Word 工作区", ""]

        # 文档列表
        lines.append(f"## 文档列表 ({len(word_docs)})")
        if word_docs:
            for doc_uri in word_docs:
                if doc_uri == active_uri:
                    lines.append(f"- ⭐ {doc_uri} (激活)")
                else:
                    lines.append(f"- {doc_uri}")
        else:
            lines.append("暂无 Word 文档连接。")

        # 激活文档详情
        if active_uri:
            # 提取文件名
            filename = active_uri.rsplit("/", 1)[-1] if "/" in active_uri else active_uri
            lines.append("")
            lines.append(f"## 激活文档: {filename}")

            # 并发拉取 stats 和 visibleContent
            stats, content = await asyncio.gather(
                self._fetch_with_timeout(active_uri, "word:get:documentStats", {"document_uri": active_uri}),
                self._fetch_with_timeout(active_uri, "word:get:visibleContent", {"document_uri": active_uri}),
            )

            # 渲染元数据
            if stats is not None:
                page_count = stats.get("pageCount", "N/A")
                word_count = stats.get("wordCount", 0)
                paragraph_count = stats.get("paragraphCount", "N/A")
                word_count_str = f"{word_count:,}" if isinstance(word_count, int) else str(word_count)
                lines.append(f"- 总页数: {page_count}")
                lines.append(f"- 总字数: {word_count_str}")
                lines.append(f"- 段落数: {paragraph_count}")
            else:
                lines.append("[元数据不可用: 请求超时]")

            # 渲染可见内容
            lines.append("")
            lines.append("## 当前可见内容")
            if content is not None:
                text = content.get("text", "")
                if text:
                    lines.append(text)
                else:
                    lines.append("(空)")
            else:
                lines.append("[可见内容不可用: 请求超时]")

        return "\n".join(lines)

    def update_from_uri(self, uri: str) -> None:
        self._priority, self._fullscreen = parse_window_uri_params(
            uri, self._priority, self._fullscreen, log_prefix="Word window resource"
        )

    # ── Internal helpers ──

    async def _fetch_with_timeout(self, document_uri: str, event: str, params: dict[str, Any]) -> dict[str, Any] | None:
        """通用 3s 超时拉取，失败返回 None."""
        try:
            response = await asyncio.wait_for(
                self.workspace.emit_to_document(document_uri, event, params),
                timeout=self.FETCH_TIMEOUT,
            )
            if response.get("success"):
                data: dict[str, Any] = response.get("data", {})
                return data
            return None
        except (TimeoutError, ValueError) as e:
            logger.warning(f"Fetch failed for {event} on {document_uri}: {e}")
            return None
