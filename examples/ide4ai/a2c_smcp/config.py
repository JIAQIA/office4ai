# filename: config.py
# @Time    : 2025/10/29 12:01
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
MCP Server 配置管理 | MCP Server Configuration Management

该模块负责管理 MCP Server 的配置，包括 IDE 实例的初始化参数
This module manages MCP Server configuration, including IDE instance initialization parameters
"""

from typing import Any, Literal

from confz import BaseConfig, CLArgSource, EnvSource
from confz.base_config import BaseConfigMetaclass
from pydantic import Field, field_validator

from ide4ai.base import WorkspaceSetting
from ide4ai.environment.terminal.command_filter import CommandFilterConfig


class MCPServerConfig(BaseConfig, metaclass=BaseConfigMetaclass):
    """
    MCP Server 配置类 | MCP Server Configuration Class

    支持从环境变量和命令行参数加载配置
    Supports loading configuration from environment variables and command-line arguments

    Attributes:
        transport: 传输模式 | Transport mode (stdio, sse, streamable-http)
        host: 服务器主机地址 | Server host address (for sse/streamable-http)
        port: 服务器端口 | Server port (for sse/streamable-http)
        cmd_white_list: 命令白名单 | Command whitelist
        root_dir: 根目录 | Root directory
        project_name: 项目名称 | Project name
        render_with_symbols: 是否渲染符号 | Whether to render symbols
        max_active_models: 最大活跃模型数 | Maximum active models
        cmd_time_out: 命令超时时间(秒) | Command timeout (seconds)
        enable_simple_view_mode: 是否启用简化视图模式 | Whether to enable simple view mode
        workspace_setting: 工作区设置 | Workspace settings

    环境变量映射 | Environment Variable Mapping:
        - TRANSPORT -> transport
        - HOST -> host
        - PORT -> port
        - PROJECT_ROOT -> root_dir
        - PROJECT_NAME -> project_name
        - CMD_WHITE_LIST -> cmd_white_list (逗号分隔 | comma separated)
        - CMD_TIMEOUT -> cmd_time_out
        - RENDER_WITH_SYMBOLS -> render_with_symbols
        - MAX_ACTIVE_MODELS -> max_active_models
        - ENABLE_SIMPLE_VIEW_MODE -> enable_simple_view_mode

    命令行参数映射 | Command-line Argument Mapping:
        - --transport -> transport
        - --host -> host
        - --port -> port
        - --root-dir -> root_dir
        - --project-name -> project_name
        - --cmd-white-list -> cmd_white_list (逗号分隔 | comma separated)
        - --cmd-timeout -> cmd_time_out
        - --render-with-symbols -> render_with_symbols
        - --max-active-models -> max_active_models
        - --enable-simple-view-mode -> enable_simple_view_mode
    """

    # 配置源优先级：命令行参数 > 环境变量
    # Configuration source priority: Command-line arguments > Environment variables
    CONFIG_SOURCES = [
        EnvSource(
            allow_all=True,
            prefix="",  # 不使用前缀，直接使用环境变量名 | No prefix, use environment variable name directly
            remap={
                "TRANSPORT": "transport",
                "HOST": "host",
                "PORT": "port",
                "PROJECT_ROOT": "root_dir",
                "PROJECT_NAME": "project_name",
                "CMD_WHITE_LIST": "cmd_white_list",
                "CMD_TIMEOUT": "cmd_time_out",
                "RENDER_WITH_SYMBOLS": "render_with_symbols",
                "MAX_ACTIVE_MODELS": "max_active_models",
                "ENABLE_SIMPLE_VIEW_MODE": "enable_simple_view_mode",
            },
        ),
        CLArgSource(
            prefix="",  # 不使用前缀 | No prefix
            remap={
                "transport": "transport",
                "host": "host",
                "port": "port",
                "root-dir": "root_dir",
                "project-name": "project_name",
                "cmd-white-list": "cmd_white_list",
                "cmd-timeout": "cmd_time_out",
                "render-with-symbols": "render_with_symbols",
                "max-active-models": "max_active_models",
                "enable-simple-view-mode": "enable_simple_view_mode",
            },
        ),
    ]

    # 传输模式配置 | Transport mode configuration
    transport: Literal["stdio", "sse", "streamable-http"] = Field(
        default="stdio",
        description="传输模式：stdio(标准输入输出), sse(Server-Sent Events), streamable-http(Streamable HTTP) | "
        "Transport mode: stdio, sse, streamable-http",
    )
    host: str = Field(
        default="127.0.0.1",
        description="服务器主机地址(仅用于 sse 和 streamable-http 模式) | Server host (only for sse and streamable-http)",
    )
    port: int = Field(
        default=8000,
        description="服务器端口(仅用于 sse 和 streamable-http 模式) | Server port (only for sse and streamable-http)",
    )

    # IDE 配置 | IDE configuration
    cmd_white_list: list[str] = Field(
        default_factory=lambda: ["ls", "pwd", "echo", "cat", "grep", "find", "head", "tail", "wc"],
        description="命令白名单，逗号分隔的字符串会被自动解析为列表 | Command whitelist, comma-separated string will be "
        "automatically parsed to list",
    )
    root_dir: str = Field(default=".", description="项目根目录 | Project root directory")
    project_name: str = Field(default="mcp-project", description="项目名称 | Project name")
    render_with_symbols: bool = Field(default=True, description="是否渲染符号 | Whether to render symbols")
    max_active_models: int = Field(default=3, description="最大活跃模型数 | Maximum active models")
    cmd_time_out: int = Field(default=10, description="命令超时时间(秒) | Command timeout (seconds)")
    enable_simple_view_mode: bool = Field(
        default=True,
        description="是否启用简化视图模式 | Whether to enable simple view mode",
    )
    workspace_setting: WorkspaceSetting | None = Field(default=None, description="工作区设置 | Workspace settings")

    @field_validator("cmd_white_list", mode="before")
    @classmethod
    def parse_cmd_white_list(cls, v: Any) -> list[str]:
        """
        解析命令白名单 | Parse command whitelist

        支持字符串（逗号分隔）和列表两种格式
        Supports both string (comma-separated) and list formats

        Args:
            v: 输入值 | Input value

        Returns:
            list[str]: 命令列表 | Command list
        """
        if isinstance(v, str):
            # 如果是字符串，按逗号分隔并去除空格 | If string, split by comma and strip spaces
            return [cmd.strip() for cmd in v.split(",") if cmd.strip()]
        elif isinstance(v, list):
            # 如果已经是列表，直接返回 | If already a list, return as is
            return v
        # 其他情况返回空列表 | Return empty list for other cases
        return []

    def to_ide_kwargs(self) -> dict[str, Any]:
        """
        转换为 IDE 初始化参数 | Convert to IDE initialization parameters

        Returns:
            dict: IDE 初始化参数字典 | IDE initialization parameters dict
        """
        # 将 cmd_white_list 转换为 CommandFilterConfig
        # Convert cmd_white_list to CommandFilterConfig
        cmd_filter = CommandFilterConfig.from_white_list(self.cmd_white_list)

        return {
            "cmd_filter": cmd_filter,
            "root_dir": self.root_dir,
            "project_name": self.project_name,
            "render_with_symbols": self.render_with_symbols,
            "max_active_models": self.max_active_models,
            "cmd_time_out": self.cmd_time_out,
            "enable_simple_view_mode": self.enable_simple_view_mode,
            "workspace_setting": self.workspace_setting,
        }
