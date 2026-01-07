"""
Socket.IO Request-Response Handler

Provides request-response mechanism for Socket.IO events.
"""

import asyncio
import logging
from typing import Any

import socketio  # type: ignore[import-untyped]

logger = logging.getLogger(__name__)

# 全局 Socket.IO 服务器引用（用于请求-响应机制）
_server_instance = None

# 全局请求表：request_id -> Future
_pending_requests: dict[str, asyncio.Future] = {}


def set_server_instance(server: socketio.AsyncServer) -> None:
    """
    Set the global Socket.IO server instance.

    Args:
        server: Socket.IO server instance
    """
    global _server_instance
    _server_instance = server


async def emit_with_response(sid: str, event: str, data: dict[str, Any], timeout: float = 10.0) -> dict[str, Any]:
    """
    发送 Socket.IO 事件并等待响应

    Args:
        sid: Session ID (目标 socket ID)
        event: 事件名称
        data: 事件数据 (必须包含 requestId)
        timeout: 超时时间（秒）

    Returns:
        dict: Add-In 返回的响应数据

    Raises:
        ValueError: 如果缺少 requestId
        TimeoutError: 如果请求超时
        RuntimeError: 如果服务器未运行
    """
    global _server_instance, _pending_requests

    if _server_instance is None:
        raise RuntimeError("Socket.IO server is not running")

    request_id = data.get("requestId")
    if not request_id:
        raise ValueError("Missing requestId in data")

    # 创建 Future 等待响应
    future: asyncio.Future[dict[str, Any]] = asyncio.Future()
    _pending_requests[request_id] = future

    try:
        # 发送事件
        logger.debug(f"Emitting {event} to {sid}, requestId={request_id}")
        _server_instance.emit(event, data, to=sid)

        # 等待响应
        result: dict[str, Any] = await asyncio.wait_for(future, timeout=timeout)
        logger.debug(f"Received response for requestId={request_id}")
        return result

    except asyncio.TimeoutError:
        logger.warning(f"Request {request_id} timed out after {timeout}s")
        del _pending_requests[request_id]
        raise TimeoutError(f"Request {request_id} timed out")

    finally:
        # 清理
        _pending_requests.pop(request_id, None)


def handle_response(request_id: str, response_data: dict[str, Any]) -> None:
    """
    处理 Add-In 返回的响应

    Args:
        request_id: 请求 ID
        response_data: 响应数据
    """
    global _pending_requests

    future = _pending_requests.get(request_id)
    if future and not future.done():
        logger.debug(f"Setting result for requestId={request_id}")
        future.set_result(response_data)
    else:
        logger.warning(f"No pending future found for requestId={request_id}")
