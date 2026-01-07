"""
Common Socket.IO DTOs

Defines base structures for Socket.IO communication between Workspace and Add-In clients.
"""

from datetime import datetime
from typing import Any, ClassVar, Optional

from pydantic import BaseModel, ConfigDict, Field


class SocketIOBaseModel(BaseModel):
    """
    Base model for all Socket.IO DTOs.

    Provides automatic snake_case ↔ camelCase conversion for protocol compliance:
    - Internal Python: snake_case (PEP 8 compliant)
    - External JSON: camelCase (Socket.IO protocol)

    Example:
        >>> class MyRequest(SocketIOBaseModel):
        ...     request_id: str = Field(alias="requestId")
        ...
        >>> # Internal: obj.request_id
        >>> # External: {"requestId": "..."}
    """

    model_config: ClassVar[ConfigDict] = ConfigDict(
        populate_by_name=True,  # Accept both alias and field name
    )


class BaseRequest(SocketIOBaseModel):
    """
    Base structure for Server → Client requests.

    When Workspace sends a command to an Add-In, all requests inherit from this.

    Uses Pydantic aliases for protocol compliance:
    - Internal: snake_case (PEP 8 compliant)
    - External: camelCase (Socket.IO protocol)
    """

    request_id: str = Field(
        ...,
        alias="requestId",
        description="Unique request identifier for matching responses",
    )
    document_uri: str = Field(
        ...,
        alias="documentUri",
        description="Document URI (file:///path/to.docx)",
    )
    timestamp: int | None = Field(
        default_factory=lambda: int(datetime.now().timestamp() * 1000),
        alias="timestamp",
        description="Client timestamp in milliseconds",
    )


class BaseResponse(SocketIOBaseModel):
    """
    Base structure for Client → Server responses.

    When Add-In responds to a Workspace command, all responses inherit from this.

    Uses Pydantic aliases for protocol compliance:
    - Internal: snake_case (PEP 8 compliant)
    - External: camelCase (Socket.IO protocol)
    """

    request_id: str = Field(
        ...,
        alias="requestId",
        description="Request ID being responded to",
    )
    success: bool = Field(..., alias="success", description="Whether the operation succeeded")
    data: dict[str, Any] | None = Field(
        default=None,
        alias="data",
        description="Response data",
    )
    error: Optional["ErrorResponse"] = Field(
        default=None,
        alias="error",
        description="Error details if failed",
    )
    timestamp: int = Field(
        ...,
        alias="timestamp",
        description="Server timestamp in milliseconds",
    )
    duration: int | None = Field(
        default=None,
        alias="duration",
        description="Operation duration in milliseconds",
    )


class ErrorResponse(SocketIOBaseModel):
    """
    Standardized error information.

    Uses Pydantic aliases for protocol compliance.
    """

    code: str = Field(
        ...,
        alias="code",
        description="Error code (e.g., '3000')",
    )
    message: str = Field(
        ...,
        alias="message",
        description="Human-readable error message",
    )
    details: dict[str, Any] | None = Field(
        default=None,
        alias="details",
        description="Additional error details",
    )


class ErrorCode:
    """
    Standard error codes for Socket.IO communication.

    Code ranges:
    - 1xxx: General errors
    - 2xxx: Authentication errors
    - 3xxx: Office API errors
    - 4xxx: Validation errors
    """

    # General errors (1xxx)
    UNKNOWN_ERROR = "1000"
    INVALID_REQUEST = "1001"
    TIMEOUT = "1002"
    NOT_IMPLEMENTED = "1003"

    # Authentication errors (2xxx)
    UNAUTHORIZED = "2000"
    TOKEN_EXPIRED = "2001"
    INVALID_TOKEN = "2002"
    HANDSHAKE_FAILED = "2003"

    # Office API errors (3xxx)
    OFFICE_API_ERROR = "3000"
    DOCUMENT_NOT_FOUND = "3001"
    SELECTION_EMPTY = "3002"
    OPERATION_FAILED = "3003"
    FILE_NOT_ACCESSIBLE = "3004"

    # Validation errors (4xxx)
    VALIDATION_ERROR = "4000"
    MISSING_PARAM = "4001"
    INVALID_PARAM = "4002"
    INVALID_PARAM_TYPE = "4003"


# Forward reference resolution
BaseResponse.model_rebuild()
