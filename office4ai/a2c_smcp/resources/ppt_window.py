"""window://office4ai/ppt Resource — PPT 文档聚合窗口资源."""

from __future__ import annotations

import asyncio
from typing import Any
from urllib.parse import parse_qs, urlencode, urlparse

from loguru import logger

from office4ai.a2c_smcp.resources.base import BaseResource, parse_window_uri_params
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import connection_manager


class PptWindowResource(BaseResource):
    """
    PPT 文档聚合窗口资源

    通过 ``window://office4ai/ppt`` 向 AI Agent 暴露 PPT 文档的实时状态，
    包括已连接文档列表、激活文档的元数据和幻灯片摘要 (±N 张)。

    每次 ``read()`` 通过 Socket.IO 拉取最新数据，3 秒超时后降级渲染。
    """

    FETCH_TIMEOUT = 3  # 秒
    DEFAULT_RANGE = 2  # ±N slides

    def __init__(self, workspace: OfficeWorkspace, priority: int = 50, fullscreen: bool = True) -> None:
        if not isinstance(priority, int) or not (0 <= priority <= 100):
            raise ValueError(f"priority must be int in [0, 100], got: {priority}")
        self.workspace = workspace
        self._priority = priority
        self._fullscreen = fullscreen
        self._range = self.DEFAULT_RANGE

    # ── BaseResource implementation ──

    @property
    def uri(self) -> str:
        query = urlencode(
            {
                "priority": str(self._priority),
                "fullscreen": "true" if self._fullscreen else "false",
            }
        )
        return f"window://office4ai/ppt?{query}"

    @property
    def base_uri(self) -> str:
        return "window://office4ai/ppt"

    @property
    def name(self) -> str:
        return "PPT 工作区"

    @property
    def description(self) -> str:
        return "PPT 文档聚合窗口，展示已连接 PPT 文档列表、激活文档的元数据和幻灯片摘要。"

    @property
    def mime_type(self) -> str:
        return "text/plain"

    async def read(self) -> str:
        clients = connection_manager.get_all_clients()
        last = self.workspace.get_last_activity()

        # 过滤 /ppt namespace 文档，按 document_uri 去重
        ppt_docs: set[str] = set()
        for c in clients:
            if c.namespace == "/ppt":
                ppt_docs.add(c.document_uri)

        # 确定激活文档
        active_uri: str | None = None
        if last is not None and last.document_uri in ppt_docs:
            active_uri = last.document_uri

        lines: list[str] = ["# PPT 工作区", ""]

        # 文档列表
        lines.append(f"## 文档列表 ({len(ppt_docs)})")
        if ppt_docs:
            for doc_uri in ppt_docs:
                if doc_uri == active_uri:
                    lines.append(f"- ⭐ {doc_uri} (激活)")
                else:
                    lines.append(f"- {doc_uri}")
        else:
            lines.append("暂无 PPT 文档连接。")

        # 激活文档详情
        if active_uri:
            filename = active_uri.rsplit("/", 1)[-1] if "/" in active_uri else active_uri
            lines.append("")
            lines.append(f"## 激活文档: {filename}")

            # 拉取 presentation 元数据 (无参 slideInfo)
            pres_info = await self._fetch_with_timeout(active_uri, "ppt:get:slideInfo", {"document_uri": active_uri})

            if pres_info is not None:
                slide_count = pres_info.get("slideCount", 0)
                current_index = pres_info.get("currentSlideIndex", 0)
                dimensions = pres_info.get("dimensions", {})
                width = dimensions.get("width", "?")
                height = dimensions.get("height", "?")
                aspect_ratio = dimensions.get("aspectRatio", "?")

                lines.append(f"- 总张数: {slide_count}")
                lines.append(f"- 尺寸: {width}×{height} pt ({aspect_ratio})")
                lines.append(f"- 当前幻灯片: 第 {current_index + 1} 张")

                # 并发拉取 ±N 张 slide 摘要
                if slide_count > 0:
                    start = max(0, current_index - self._range)
                    end = min(slide_count - 1, current_index + self._range)

                    lines.append("")
                    lines.append(f"## 幻灯片摘要 (第 {start + 1}-{end + 1} 张)")

                    tasks = [
                        self._fetch_with_timeout(
                            active_uri,
                            "ppt:get:slideInfo",
                            {"document_uri": active_uri, "slide_index": i},
                        )
                        for i in range(start, end + 1)
                    ]
                    slide_results = await asyncio.gather(*tasks)

                    for idx, slide_data in enumerate(slide_results):
                        i = start + idx
                        is_current = i == current_index
                        marker = "➡️ " if is_current else ""
                        current_label = " (当前)" if is_current else ""

                        if slide_data is not None:
                            slide_info = slide_data.get("slideInfo", {})
                            title = slide_info.get("title", f"幻灯片 {i + 1}")
                            elements = slide_data.get("elements", [])
                            notes = slide_info.get("notes", "")

                            lines.append("")
                            lines.append(f"### {marker}第 {i + 1} 张: {title}{current_label}")

                            # 元素计数（按类型）
                            if elements:
                                type_counts: dict[str, int] = {}
                                for elem in elements:
                                    elem_type = elem.get("type", "未知")
                                    type_counts[elem_type] = type_counts.get(elem_type, 0) + 1
                                elem_str = ", ".join(f"{t}×{c}" for t, c in type_counts.items())
                                lines.append(f"- 元素: {elem_str}")
                            else:
                                lines.append("- 元素: (无)")

                            # 备注
                            lines.append(f"- 备注: {notes if notes else '(无)'}")
                        else:
                            lines.append("")
                            lines.append(f"### {marker}第 {i + 1} 张{current_label}")
                            lines.append("[幻灯片信息不可用: 请求超时]")
            else:
                lines.append("[元数据不可用: 请求超时]")

        return "\n".join(lines)

    def update_from_uri(self, uri: str) -> None:
        self._priority, self._fullscreen = parse_window_uri_params(
            uri, self._priority, self._fullscreen, log_prefix="PPT window resource"
        )

        # PPT-specific: range parameter
        parsed = urlparse(uri)
        params = parse_qs(parsed.query)
        if "range" in params:
            try:
                new_range = int(params["range"][0])
                if 0 <= new_range <= 10:
                    if new_range != self._range:
                        logger.debug(f"PPT window resource range: {self._range} -> {new_range}")
                        self._range = new_range
                else:
                    logger.warning(f"Invalid range value in URI: {new_range}, must be in [0, 10]")
            except (ValueError, IndexError) as e:
                logger.warning(f"Failed to parse range from URI: {e}")

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
