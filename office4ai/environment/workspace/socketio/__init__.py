"""
Workspace Socket.IO Server

This module provides Socket.IO server functionality for communication
between office4ai Workspace and Office Add-In clients.

Architecture:
- Server binds to 127.0.0.1 (localhost only)
- Uses simple handshake (clientId + documentUri)
- Manages documentUri → socketId mappings
- Provides /word, /ppt, /excel namespaces
"""

from office4ai.environment.workspace.socketio.server import create_socketio_server

__all__ = [
    "create_socketio_server",
]
