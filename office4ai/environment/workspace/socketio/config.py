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
        "http://localhost:3000",
        "http://127.0.0.1:3000",
        "http://localhost:*",  # Allow any localhost port
        "capacitor://localhost",  # For Electron/Capacitor
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
