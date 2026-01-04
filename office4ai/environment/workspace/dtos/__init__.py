"""
Workspace Socket.IO DTOs

This package contains data transfer objects for Socket.IO communication
between office4ai (Workspace) and Office Add-In clients.

Architecture:
- These DTOs are internal to the Workspace environment
- They define the protocol for Add-In ↔ Workspace communication
- NOT to be confused with office4ai/dtos/ which are for A2C-SMCP protocol
"""

from office4ai.environment.workspace.dtos.common import (
    BaseRequest,
    BaseResponse,
    ErrorCode,
    ErrorResponse,
)

__all__ = [
    "BaseRequest",
    "BaseResponse",
    "ErrorResponse",
    "ErrorCode",
]
