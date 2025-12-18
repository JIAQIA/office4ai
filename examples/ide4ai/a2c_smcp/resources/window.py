# filename: window.py
# @Time    : 2025/11/04 16:48
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
窗口资源实现 | Window Resource Implementation

提供 IDE 窗口内容的 MCP 资源封装
Provides MCP resource wrapper for IDE window content

遵循 window:// 协议规范：
Follows window:// protocol specification:
- URI 格式 | URI format: window://{{project_name}}?priority={{priority}}&fullscreen={{fullscreen}}
- priority: 0-100 的整数，影响布局排序 | integer 0-100, affects layout ordering
- fullscreen: 布尔值，是否全屏显示 | boolean, whether to display fullscreen
"""

from urllib.parse import parse_qs, urlencode, urlparse

from loguru import logger

from ide4ai.a2c_smcp.resources.base import BaseResource
from ide4ai.base import IDE


class WindowResource(BaseResource):
    """
    窗口资源 | Window Resource

    将 IDE 的渲染内容封装为 MCP 资源
    Wraps IDE render content as MCP resource
    """

    def __init__(self, ide: IDE, priority: int = 0, fullscreen: bool = True) -> None:
        """
        初始化窗口资源 | Initialize window resource

        Args:
            ide: IDE 实例 | IDE instance
            priority: 窗口优先级（0-100），默认 0 | Window priority (0-100), default 0
            fullscreen: 是否全屏显示，默认 True | Whether to display fullscreen, default True
        """
        super().__init__(ide)
        if not isinstance(priority, int) or not (0 <= priority <= 100):
            raise ValueError(f"priority must be int in [0, 100], got: {priority}")
        self._priority = priority
        self._fullscreen = fullscreen

    @property
    def uri(self) -> str:
        """
        资源 URI | Resource URI

        Returns:
            str: window:// 协议的 URI | URI with window:// protocol
        """
        # 构建查询参数 | Build query parameters
        query_params = {
            "priority": str(self._priority),
            "fullscreen": "true" if self._fullscreen else "false",
        }
        query_string = urlencode(query_params)

        # 返回符合 window:// 协议规范的 URI | Return URI following window:// protocol
        return f"window://{self.ide.project_name}?{query_string}"

    @property
    def name(self) -> str:
        """
        资源名称 | Resource name

        Returns:
            str: 资源名称 | Resource name
        """
        return f"IDE Window - {self.ide.project_name}"

    @property
    def description(self) -> str:
        """
        资源描述 | Resource description

        Returns:
            str: 资源描述 | Resource description
        """
        return "IDE 窗口内容，包含工作区和终端信息 | IDE window content including workspace and terminal information"

    @property
    def mime_type(self) -> str:
        """
        资源 MIME 类型 | Resource MIME type

        Returns:
            str: MIME 类型 | MIME type
        """
        return "text/plain"

    async def read(self) -> str:
        """
        读取资源内容 | Read resource content

        使用 IDE 的 render 方法获取当前窗口内容
        Use IDE's render method to get current window content

        Returns:
            str: IDE 渲染的窗口内容 | IDE rendered window content
        """
        return self.ide.render()

    def update_from_uri(self, uri: str) -> None:
        """
        从 URI 更新窗口参数 | Update window parameters from URI

        解析 URI 中的 priority 和 fullscreen 参数，并更新到当前实例
        Parse priority and fullscreen parameters from URI and update current instance

        Args:
            uri: 包含新参数的完整 URI | Full URI with new parameters
        """
        parsed = urlparse(uri)
        query_params = parse_qs(parsed.query)

        # 更新 priority 参数 | Update priority parameter
        if "priority" in query_params:
            try:
                new_priority = int(query_params["priority"][0])
                if 0 <= new_priority <= 100:
                    if new_priority != self._priority:
                        logger.debug(
                            f"更新窗口资源 priority | Update window resource priority: {self._priority} -> {new_priority}",
                        )
                        self._priority = new_priority
                else:
                    logger.warning(f"Invalid priority value in URI: {new_priority}, must be in [0, 100]")
            except (ValueError, IndexError) as e:
                logger.warning(f"Failed to parse priority from URI: {e}")
        # 更新 fullscreen 参数 | Update fullscreen parameter
        if "fullscreen" in query_params:
            try:
                fullscreen_str = query_params["fullscreen"][0].lower()
                # 验证是否为有效的布尔值 | Validate if it's a valid boolean value
                new_fullscreen = None
                if fullscreen_str in {"true", "1", "yes", "on"}:
                    new_fullscreen = True
                elif fullscreen_str in {"false", "0", "no", "off"}:
                    new_fullscreen = False
                else:
                    # 无效值，记录警告并忽略 | Invalid value, log warning and ignore
                    logger.warning(f"Invalid fullscreen value in URI: {fullscreen_str}, ignoring")

                # 只有当解析成功且值确实改变时才更新 | Only update if parsed successfully and value changed
                if new_fullscreen is not None and new_fullscreen != self._fullscreen:
                    logger.debug(
                        f"更新窗口资源 fullscreen | Update window resource fullscreen: {self._fullscreen} -> {new_fullscreen}",
                    )
                    self._fullscreen = new_fullscreen
            except (IndexError, AttributeError) as e:
                logger.warning(f"Failed to parse fullscreen from URI: {e}")
