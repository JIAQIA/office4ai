# filename: server.py
# @Time    : 2025/11/03 20:28
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
MCP Server 基类 | MCP Server Base Class

提供通用的 MCP 协议服务器实现，支持多语言 IDE 扩展
Provides generic MCP protocol server implementation, supporting multi-language IDE extensions
"""

from abc import ABC, abstractmethod
from typing import Any

from loguru import logger
from mcp.server import Server
from mcp.server.sse import SseServerTransport
from mcp.server.stdio import stdio_server
from mcp.server.streamable_http_manager import StreamableHTTPSessionManager
from mcp.types import Resource, Tool
from pydantic import AnyUrl
from starlette.applications import Starlette
from starlette.responses import Response
from starlette.routing import Mount, Route

from ide4ai.a2c_smcp.config import MCPServerConfig
from ide4ai.a2c_smcp.resources.base import BaseResource
from ide4ai.a2c_smcp.tools.base import BaseTool
from ide4ai.base import IDE


class BaseMCPServer(ABC):
    """
    MCP Server 基类 | MCP Server Base Class

    封装 IDE 为 MCP Server，提供工具和资源
    Wraps IDE as MCP Server, providing tools and resources

    所有语言的 MCP Server 都应该继承此类
    All language-specific MCP Servers should inherit from this class
    """

    def __init__(self, config: MCPServerConfig, server_name: str) -> None:
        """
        初始化 MCP Server | Initialize MCP Server

        Args:
            config: MCP Server 配置 | MCP Server configuration
            server_name: 服务器名称 | Server name
        """
        self.config = config
        self.server = Server(server_name)

        # 初始化 IDE 实例（由子类实现）| Initialize IDE instance (implemented by subclass)
        self.ide: IDE = self._create_ide_instance()

        # 初始化工具列表 | Initialize tools list
        self.tools: dict[str, BaseTool] = {}

        # 初始化资源列表 | Initialize resources list
        self.resources: dict[str, BaseResource] = {}

        # 注册工具 | Register tools
        self._register_tools()

        # 注册资源 | Register resources
        self._register_resources()

        # 设置 MCP 处理器 | Setup MCP handlers
        self._setup_handlers()

        logger.info(
            f"MCP Server 初始化完成 | MCP Server initialized: server={server_name}, project={config.project_name}, root={config.root_dir}",
        )

    @abstractmethod
    def _create_ide_instance(self) -> IDE:
        """
        创建 IDE 实例 | Create IDE instance

        由子类实现，根据具体的 IDE 类型创建实例
        Implemented by subclass to create instance based on specific IDE type

        Returns:
            IDE: IDE 实例 | IDE instance
        """
        pass

    @abstractmethod
    def _register_tools(self) -> None:
        """
        注册所有工具 | Register all tools

        由子类实现，注册特定语言的工具
        Implemented by subclass to register language-specific tools
        """
        pass

    @abstractmethod
    def _register_resources(self) -> None:
        """
        注册所有资源 | Register all resources

        由子类实现，注册特定语言的资源
        Implemented by subclass to register language-specific resources
        """
        pass

    def close(self) -> None:
        """
        关闭 MCP Server 并清理资源 | Close MCP Server and cleanup resources

        释放 IDE 和 workspace 资源，关闭 LSP 进程
        Release IDE and workspace resources, close LSP process
        """
        try:
            if self.ide:
                # 调用 IDE 的 close 方法，由 IDE 负责清理其内部资源
                # Call IDE's close method, let IDE handle its internal resource cleanup
                self.ide.close()
                logger.info(
                    f"MCP Server 资源已清理 | MCP Server resources cleaned up: project={self.config.project_name}",
                )
        except Exception as e:
            logger.error(f"关闭 MCP Server 时出错 | Error closing MCP Server: {e}")

    def __del__(self) -> None:
        """
        析构函数，确保资源被清理 | Destructor to ensure resources are cleaned up
        """
        try:
            self.close()
        except Exception as e:
            # 在析构函数中捕获所有异常，避免影响垃圾回收
            # Catch all exceptions in destructor to avoid affecting garbage collection
            logger.error(f"析构时关闭环境出错 / Error closing environment in destructor: {e}")

    def _setup_handlers(self) -> None:
        """
        设置 MCP 协议处理器 | Setup MCP protocol handlers
        """

        @self.server.list_tools()  # type: ignore[no-untyped-call]
        async def list_tools() -> list[Tool]:
            """
            列出所有可用工具 | List all available tools

            Returns:
                list[Tool]: 工具列表 | List of tools
            """
            tools = []
            for tool in self.tools.values():
                tools.append(
                    Tool(
                        name=tool.name,
                        description=tool.description,
                        inputSchema=tool.input_schema,
                    ),
                )

            logger.debug(f"列出工具 | Listed tools: {[t.name for t in tools]}")
            return tools

        @self.server.call_tool()
        async def call_tool(name: str, arguments: dict[str, Any]) -> list[dict[str, Any]]:
            """
            调用工具 | Call tool

            Args:
                name: 工具名称 | Tool name
                arguments: 工具参数 | Tool arguments

            Returns:
                list[dict]: 工具执行结果 | Tool execution result
            """
            logger.info(f"调用工具 | Calling tool: {name}")

            # 查找工具 | Find tool
            tool = self.tools.get(name)
            if not tool:
                error_msg = f"未找到工具 | Tool not found: {name}"
                logger.error(error_msg)
                return [{"type": "text", "text": error_msg}]

            try:
                # 执行工具 | Execute tool
                result = await tool.execute(arguments)

                # 格式化返回结果 | Format return result
                return [
                    {
                        "type": "text",
                        "text": str(result),
                    },
                ]

            except Exception as e:
                error_msg = f"工具执行失败 | Tool execution failed: {e}"
                logger.exception(error_msg)
                return [{"type": "text", "text": error_msg}]

        @self.server.list_resources()  # type: ignore[no-untyped-call]
        async def list_resources() -> list[Resource]:
            """
            列出所有可用资源 | List all available resources

            Returns:
                list[Resource]: 资源列表 | List of resources
            """
            resources = []
            for resource in self.resources.values():
                resources.append(
                    Resource(
                        uri=AnyUrl(resource.uri),
                        name=resource.name,
                        description=resource.description,
                        mimeType=resource.mime_type,
                    ),
                )

            logger.debug(f"列出资源 | Listed resources: {[r.uri for r in resources]}")
            return resources

        @self.server.read_resource()  # type: ignore[no-untyped-call]
        async def read_resource(uri: str) -> str:
            """
            读取资源内容 | Read resource content

            支持动态参数更新：如果用户请求的 URI 包含不同的查询参数，
            会先调用资源的 update_from_uri 方法更新参数，然后再读取内容。
            Supports dynamic parameter updates: if the requested URI contains different query parameters,
            it will call the resource's update_from_uri method to update parameters before reading.

            Args:
                uri: 资源 URI（可能包含查询参数）| Resource URI (may contain query parameters)

            Returns:
                str: 资源内容 | Resource content
            """
            logger.info(f"读取资源 | Reading resource: {uri}")

            # 解析 base_uri（不含查询参数）用于查找资源
            # Parse base_uri (without query parameters) to find resource
            from urllib.parse import urlparse

            parsed = urlparse(uri)
            base_uri = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"

            # 使用 base_uri 查找资源 | Find resource using base_uri
            resource = self.resources.get(base_uri)
            if not resource:
                error_msg = f"未找到资源 | Resource not found: {base_uri}"
                logger.error(error_msg)
                raise ValueError(error_msg)

            try:
                # 如果请求的 URI 与资源当前 URI 不同，更新资源参数
                # If requested URI differs from resource's current URI, update parameters
                if uri != resource.uri:
                    logger.debug(f"更新资源参数 | Updating resource parameters from URI: {uri}")
                    resource.update_from_uri(uri)

                # 读取资源内容 | Read resource content
                content = await resource.read()

                # 返回文本内容 | Return text content
                return content

            except Exception as e:
                error_msg = f"资源读取失败 | Resource read failed: {e}"
                logger.exception(error_msg)
                raise ValueError(error_msg) from e

    async def run(self) -> None:
        """
        运行 MCP Server | Run MCP Server

        根据配置选择传输协议：stdio, sse 或 streamable-http
        Choose transport protocol based on configuration: stdio, sse or streamable-http
        """
        transport = self.config.transport
        logger.info(f"启动 MCP Server | Starting MCP Server with transport: {transport}")

        if transport == "stdio":
            await self._run_stdio()
        elif transport == "sse":
            await self._run_sse()
        elif transport == "streamable-http":
            await self._run_streamable_http()
        else:
            raise ValueError(f"不支持的传输模式 | Unsupported transport mode: {transport}")

    async def _run_stdio(self) -> None:
        """
        使用 stdio 传输运行 | Run with stdio transport
        """
        logger.info("使用 stdio 传输模式 | Using stdio transport")
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options(),
            )

    async def _run_sse(self) -> None:
        """
        使用 SSE 传输运行 | Run with SSE transport

        SSE (Server-Sent Events) 适用于需要服务器主动推送的场景
        SSE is suitable for scenarios requiring server-initiated push
        """
        import uvicorn

        logger.info(
            f"使用 SSE 传输模式 | Using SSE transport: http://{self.config.host}:{self.config.port}",
        )

        # 创建 SSE 传输 | Create SSE transport
        sse = SseServerTransport("/messages/")

        # 定义 SSE 处理器 | Define SSE handler
        async def handle_sse(request: Any) -> Response:
            async with sse.connect_sse(request.scope, request.receive, request._send) as streams:
                await self.server.run(
                    streams[0],
                    streams[1],
                    self.server.create_initialization_options(),
                )
            return Response()

        # 创建路由 | Create routes
        routes = [
            Route("/sse", endpoint=handle_sse, methods=["GET"]),
            Mount("/messages/", app=sse.handle_post_message),
        ]

        # 创建并运行 Starlette 应用 | Create and run Starlette app
        app = Starlette(routes=routes)
        config = uvicorn.Config(
            app,
            host=self.config.host,
            port=self.config.port,
            log_level="info",
        )
        server = uvicorn.Server(config)
        await server.serve()

    async def _run_streamable_http(self) -> None:
        """
        使用 Streamable HTTP 传输运行 | Run with Streamable HTTP transport

        Streamable HTTP 支持双向通信和流式响应，适合复杂交互场景
        Streamable HTTP supports bidirectional communication and streaming responses

        参考官方示例实现 | Reference official example implementation:
        - tests/shared/test_streamable_http.py::create_app
        """
        import uvicorn

        logger.info(
            f"使用 Streamable HTTP 传输模式 | Using Streamable HTTP transport: http://{self.config.host}:{self.config.port}/mcp",
        )

        # 创建会话管理器 | Create session manager
        # stateless=True: 每个请求都是独立的，不维护会话状态
        # stateless=True: Each request is independent, no session state is maintained
        session_manager = StreamableHTTPSessionManager(
            app=self.server,
            stateless=True,  # 无状态模式，适合简单场景 | Stateless mode, suitable for simple scenarios
            json_response=False,  # 使用 SSE 流式响应 | Use SSE streaming response
        )

        # 创建 ASGI 应用包装器类 | Create ASGI app wrapper class
        # 参考 FastMCP 的 StreamableHTTPASGIApp 实现
        # Reference FastMCP's StreamableHTTPASGIApp implementation
        class StreamableHTTPASGIApp:
            """ASGI 应用包装器 | ASGI application wrapper"""

            def __init__(self, manager: StreamableHTTPSessionManager) -> None:
                self.session_manager = manager

            async def __call__(self, scope: dict, receive: Any, send: Any) -> None:
                await self.session_manager.handle_request(scope, receive, send)

        # 创建 ASGI 应用实例 | Create ASGI app instance
        streamable_http_app = StreamableHTTPASGIApp(session_manager)

        # 创建 Starlette 应用 | Create Starlette application
        # 参考 FastMCP 实现：使用 Route 配合 ASGI 应用类
        # Reference FastMCP implementation: Use Route with ASGI app class
        app = Starlette(
            debug=False,
            routes=[
                # 使用 Route 处理 /mcp 端点（官方标准端点）
                # Use Route to handle /mcp endpoint (official standard endpoint)
                Route(
                    "/mcp",
                    endpoint=streamable_http_app,
                    methods=["GET", "POST", "DELETE"],
                ),
            ],
            lifespan=lambda app: session_manager.run(),
        )

        # 启动 HTTP 服务器 | Start HTTP server
        config = uvicorn.Config(
            app=app,
            host=self.config.host,
            port=self.config.port,
            log_level="info",
        )
        server = uvicorn.Server(config)
        await server.serve()
