"""
Socket.IO Server

Main entry point for Socket.IO server in Workspace environment.
Binds to 127.0.0.1 for localhost-only connections.
Supports both HTTP and HTTPS.
"""

import asyncio
import logging
import ssl
from pathlib import Path
from typing import Any

import socketio  # type: ignore[import-untyped]
from aiohttp import web

from .config import SocketIOConfig, default_config
from .namespaces.word import WordNamespace
from .request_handler import set_server_instance
from .services.connection_manager import connection_manager

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


def create_socketio_server(config: SocketIOConfig = default_config) -> socketio.AsyncServer:
    """
    Create and configure Socket.IO server.

    Args:
        config: Server configuration

    Returns:
        Configured Socket.IO server instance
    """
    # Create Socket.IO server
    sio = socketio.AsyncServer(
        async_mode="aiohttp",
        cors_allowed_origins=config.cors_allowed_origins,
        ping_timeout=config.ping_timeout,
        ping_interval=config.ping_interval,
        max_http_buffer_size=config.max_http_buffer_size,
        logger=config.logger,
        engineio_logger=config.engineio_logger,
    )

    # Register namespaces
    word_namespace = WordNamespace()
    sio.register_namespace(word_namespace)

    # TODO: Register PPT and Excel namespaces in future phases
    # from .namespaces.ppt import PptNamespace
    # from .namespaces.excel import ExcelNamespace
    # sio.register_namespace(PptNamespace())
    # sio.register_namespace(ExcelNamespace())

    # 保存全局引用（用于请求-响应机制）
    set_server_instance(sio)

    # Log startup
    logger.info("Socket.IO server created")
    logger.info(f"Namespaces: {', '.join(config.namespaces)}")
    logger.info(f"CORS origins: {config.cors_allowed_origins}")

    return sio


async def start_server(
    host: str = "127.0.0.1",
    port: int = 3000,
    use_https: bool = True,
    config: SocketIOConfig = default_config,
) -> None:
    """
    Start Socket.IO server with optional HTTPS support.

    Args:
        host: Host to bind to (default: 127.0.0.1 for localhost only)
        port: HTTP port to bind to (default: 3000)
        use_https: Enable HTTPS on port+1443 (default: True)
        config: Server configuration

    Example:
        >>> import asyncio
        >>> from office4ai.environment.workspace.socketio import start_server
        >>> asyncio.run(start_server())
        >>> asyncio.run(start_server(use_https=False))  # HTTP only
    """
    # Create Socket.IO server
    sio = create_socketio_server(config)

    # Create aiohttp app
    app = web.Application()
    sio.attach(app)

    # Setup routes (health check)
    async def health_check(request: Any) -> web.Response:
        return web.json_response(
            {
                "status": "ok",
                "service": "office4ai-workspace-socketio",
                "connections": connection_manager.get_connection_count(),
                "documents": connection_manager.get_document_count(),
            }
        )

    app.router.add_get("/health", health_check)

    # Create runner
    runner = web.AppRunner(app)
    await runner.setup()

    # Start HTTP server
    site_http = web.TCPSite(runner, host, port)
    await site_http.start()

    logger.info("=" * 60)
    logger.info(f"✅ Office Workspace started")
    logger.info(f"HTTP:  http://{host}:{port}")

    # Start HTTPS server if enabled
    if use_https:
        # Get certificate paths (relative to project root)
        project_root = Path(__file__).parent.parent.parent.parent.parent
        cert_path = project_root / "certs" / "cert.pem"
        key_path = project_root / "certs" / "key.pem"

        if not cert_path.exists() or not key_path.exists():
            logger.warning("⚠️  SSL certificates not found, skipping HTTPS")
            logger.warning(f"   Expected: {cert_path}")
            logger.warning(f"   Expected: {key_path}")
            logger.warning("   Run: openssl req -x509 -newkey rsa:4096 -keyout certs/key.pem -out certs/cert.pem -days 365 -nodes")
        else:
            # Create SSL context
            ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            ssl_context.load_cert_chain(cert_path, key_path)

            # Start HTTPS site on port + 1443 (3000 + 1443 = 4443)
            https_port = port + 1443
            site_https = web.TCPSite(runner, host, https_port, ssl_context=ssl_context)
            await site_https.start()

            logger.info(f"HTTPS: https://{host}:{https_port}")

    logger.info("=" * 60)
    logger.info(f"Health check: http://{host}:{port}/health")
    logger.info(f"Bind address: {host} (localhost only)")
    logger.info("=" * 60)

    # Keep server running
    try:
        # In production, this would be part of the Workspace lifecycle
        # For now, just keep the task alive
        import asyncio

        while True:
            await asyncio.sleep(3600)  # Sleep 1 hour

    except asyncio.CancelledError:
        logger.info("Server shutdown requested")
    finally:
        await runner.cleanup()
        logger.info("Server stopped")


# ============================================================================
# Convenience Functions
# ============================================================================


def create_app(config: SocketIOConfig = default_config) -> web.Application:
    """
    Create aiohttp app with Socket.IO attached.

    Useful for integrating with existing aiohttp applications.

    Args:
        config: Server configuration

    Returns:
        aiohttp Application with Socket.IO attached

    Example:
        >>> from office4ai.environment.workspace.socketio import create_app
        >>> app = create_app()
        >>> web.run_app(app, host="127.0.0.1", port=3000)
    """
    sio = create_socketio_server(config)
    app = web.Application()
    sio.attach(app)

    # Add health check
    async def health_check(request: Any) -> web.Response:
        return web.json_response(
            {
                "status": "ok",
                "service": "office4ai-workspace-socketio",
                "connections": connection_manager.get_connection_count(),
                "documents": connection_manager.get_document_count(),
            }
        )

    app.router.add_get("/health", health_check)

    return app


if __name__ == "__main__":
    # Run server directly (for testing)
    import asyncio
    import argparse

    parser = argparse.ArgumentParser(description="Office Workspace Socket.IO Server")
    parser.add_argument(
        "--host",
        type=str,
        default="127.0.0.1",
        help="Host to bind to (default: 127.0.0.1)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=3000,
        help="HTTP port to bind to (default: 3000)",
    )
    parser.add_argument(
        "--use-https",
        type=str,
        default="true",
        choices=["true", "false"],
        help="Enable HTTPS on port+1443 (default: true)",
    )

    args = parser.parse_args()
    use_https = args.use_https.lower() == "true"

    asyncio.run(start_server(host=args.host, port=args.port, use_https=use_https))
