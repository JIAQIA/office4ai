# filename: server.py
# @Time    : 2025/10/29 12:01
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Python IDE MCP Server 实现 | Python IDE MCP Server Implementation

继承通用 MCP Server 基类，为 Python IDE 提供特定实现
Inherits from generic MCP Server base class, providing specific implementation for Python IDE
"""

from loguru import logger

from ide4ai.a2c_smcp.config import MCPServerConfig
from ide4ai.a2c_smcp.resources import WindowResource
from ide4ai.a2c_smcp.server import BaseMCPServer
from ide4ai.a2c_smcp.tools import BashTool, EditTool, GlobTool, GrepTool, ReadTool, WriteTool
from ide4ai.base import IDE
from ide4ai.ides import PyIDESingleton


class PythonIDEMCPServer(BaseMCPServer):
    """
    Python IDE MCP Server

    继承通用 MCP Server 基类，封装 PythonIDE 的能力
    Inherits from generic MCP Server base class, wrapping PythonIDE capabilities
    """

    def __init__(self, config: MCPServerConfig) -> None:
        """
        初始化 Python IDE MCP Server | Initialize Python IDE MCP Server

        Args:
            config: MCP Server 配置 | MCP Server configuration
        """
        # 调用父类初始化 | Call parent class initialization
        super().__init__(config, server_name="python-ide-mcp")

    def _create_ide_instance(self) -> IDE:
        """
        创建 Python IDE 实例 | Create Python IDE instance

        使用 PyIDESingleton 获取 IDE 实例
        Get IDE instance using PyIDESingleton

        Returns:
            IDE: Python IDE 实例 | Python IDE instance
        """
        ide_singleton = PyIDESingleton(**self.config.to_ide_kwargs())
        return ide_singleton.ide

    def _register_tools(self) -> None:
        """
        注册所有工具 | Register all tools
        """
        # 注册 Bash 工具 | Register Bash tool
        bash_tool = BashTool(self.ide)
        self.tools[bash_tool.name] = bash_tool

        # 注册 Glob 工具 | Register Glob tool
        glob_tool = GlobTool(self.ide)
        self.tools[glob_tool.name] = glob_tool

        # 注册 Grep 工具 | Register Grep tool
        grep_tool = GrepTool(self.ide)
        self.tools[grep_tool.name] = grep_tool

        # 注册 Read 工具 | Register Read tool
        read_tool = ReadTool(self.ide)
        self.tools[read_tool.name] = read_tool

        # 注册 Edit 工具 | Register Edit tool
        edit_tool = EditTool(self.ide)
        self.tools[edit_tool.name] = edit_tool

        # 注册 Write 工具 | Register Write tool
        write_tool = WriteTool(self.ide)
        self.tools[write_tool.name] = write_tool

        logger.info(f"已注册工具 | Registered tools: {list(self.tools.keys())}")

    def _register_resources(self) -> None:
        """
        注册所有资源 | Register all resources
        """
        # 注册窗口资源 | Register Window resource
        # 使用 base_uri 作为 key，确保相同资源不同参数使用同一个实例
        # Use base_uri as key to ensure same resource with different params uses same instance
        window_resource = WindowResource(self.ide, priority=0, fullscreen=True)
        self.resources[window_resource.base_uri] = window_resource

        logger.info(f"已注册资源 | Registered resources: {list(self.resources.keys())}")

        # TODO: 注册其他资源 | Register other resources


async def async_main() -> None:
    """
    异步主函数 | Async main function

    使用 confz 从环境变量和命令行参数读取配置并启动 MCP Server
    Use confz to read configuration from environment variables and command-line arguments, then start MCP Server

    配置优先级 | Configuration Priority:
        命令行参数 > 环境变量 > 默认值
        Command-line arguments > Environment variables > Default values

    环境变量 | Environment Variables:
        - TRANSPORT: 传输模式 | Transport mode (default: "stdio")
        - HOST: 服务器主机地址 | Server host (default: "127.0.0.1")
        - PORT: 服务器端口 | Server port (default: 8000)
        - PROJECT_ROOT: 项目根目录 | Project root directory (default: ".")
        - PROJECT_NAME: 项目名称 | Project name (default: "mcp-project")
        - CMD_WHITE_LIST: 命令白名单，逗号分隔 | Command whitelist, comma separated
        - CMD_TIMEOUT: 命令超时时间(秒) | Command timeout in seconds (default: 10)
        - RENDER_WITH_SYMBOLS: 是否渲染符号 | Whether to render symbols (default: true)
        - MAX_ACTIVE_MODELS: 最大活跃模型数 | Maximum active models (default: 3)
        - ENABLE_SIMPLE_VIEW_MODE: 是否启用简化视图模式 | Whether to enable simple view mode (default: true)

    命令行参数 | Command-line Arguments:
        - --transport: 传输模式 | Transport mode
        - --host: 服务器主机地址 | Server host
        - --port: 服务器端口 | Server port
        - --root-dir: 项目根目录 | Project root directory
        - --project-name: 项目名称 | Project name
        - --cmd-white-list: 命令白名单，逗号分隔 | Command whitelist, comma separated
        - --cmd-timeout: 命令超时时间(秒) | Command timeout in seconds
        - --render-with-symbols: 是否渲染符号 | Whether to render symbols
        - --max-active-models: 最大活跃模型数 | Maximum active models
        - --enable-simple-view-mode: 是否启用简化视图模式 | Whether to enable simple view mode
    """
    # 使用 confz 加载配置 | Load configuration using confz
    # confz 会自动从环境变量和命令行参数中读取配置
    # confz will automatically read configuration from environment variables and command-line arguments
    config = MCPServerConfig()

    logger.info(
        f"启动 MCP Server | Starting MCP Server: "
        f"transport={config.transport}, "
        f"host={config.host}, "
        f"port={config.port}, "
        f"root_dir={config.root_dir}, "
        f"project_name={config.project_name}, "
        f"cmd_white_list={config.cmd_white_list}, "
        f"cmd_timeout={config.cmd_time_out}, "
        f"render_with_symbols={config.render_with_symbols}, "
        f"max_active_models={config.max_active_models}, "
        f"enable_simple_view_mode={config.enable_simple_view_mode}",
    )

    # 创建并运行 server | Create and run server
    server = PythonIDEMCPServer(config)
    await server.run()


def main() -> None:
    """
    同步入口函数 | Synchronous entry point

    用于命令行调用，内部使用 asyncio.run() 运行异步主函数
    For command-line invocation, internally uses asyncio.run() to run the async main function
    """
    import asyncio

    asyncio.run(async_main())
