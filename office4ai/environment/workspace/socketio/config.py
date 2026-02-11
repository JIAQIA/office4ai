"""
Socket.IO Server Configuration
"""

from pydantic import BaseModel


class SocketIOConfig(BaseModel):
    """
    Socket.IO server configuration.
    """

    # Server binding
    host: str = "127.0.0.1"  # ⭐ Only bind to localhost
    port: int = 3000

    # CORS settings (only allow localhost)
    cors_allowed_origins: list[str] = [
        # Python Socket.IO server (HTTP)
        "http://localhost:3000",
        "http://127.0.0.1:3000",
        # Excel Add-In (HTTPS)
        "https://localhost:3001",
        "https://127.0.0.1:3001",
        # Word Add-In (HTTPS) - 开发端口
        "https://localhost:3002",
        "https://127.0.0.1:3002",
        # Word Add-In (HTTPS) - E2E 测试端口 (4443 = 3000 + 1443)
        "https://localhost:4443",
        "https://127.0.0.1:4443",
        # PowerPoint Add-In (HTTPS)
        "https://localhost:3003",
        "https://127.0.0.1:3003",
        # Capacitor/Electron
        "capacitor://localhost",
        # 线上部署的 Add-In taskpane (连接本地 Socket.IO 服务器)
        "https://office4ai.turingfocus.cn",
    ]

    # Engine.IO settings
    ping_timeout: int = 60000  # 60 seconds
    ping_interval: int = 25000  # 25 seconds
    max_http_buffer_size: int = 1000000  # 1MB

    # Logging
    logger: bool = True  # Socket.IO logger
    engineio_logger: bool = False  # Engine.IO logger (too verbose)

    # Connection settings
    reconnection: bool = True
    reconnection_attempts: int = 10
    reconnection_delay: int = 1000  # milliseconds

    # Timeouts
    request_timeout: int = 30000  # 30 seconds for client to respond

    # Namespaces
    namespaces: list[str] = ["/word", "/ppt", "/excel"]


# Default configuration instance
default_config = SocketIOConfig()
