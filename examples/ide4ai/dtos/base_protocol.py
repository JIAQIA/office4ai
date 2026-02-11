# filename: base_protocol.py
# @Time    : 2024/4/29 14:03
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
from enum import Enum
from typing import Literal

from pydantic import BaseModel


class LSPMessage(BaseModel):
    """
    LSPMessage
    """

    jsonrpc: Literal["2.0"] = "2.0"


class LSPRequestMessage(LSPMessage):
    """
    LSPRequestMessage
    """

    # The request id.
    id: int | str
    # The method to be invoked.
    method: str
    # The method's params.
    params: list | dict | None = None


class LSPResponseError(BaseModel):
    """
    LSPResponseError
    """

    # A number indicating the error type that occurred.
    code: int
    # A string providing a short description of the error.
    message: str
    # A Primitive or Structured value that contains additional information about the error. Can be omitted.
    data: str | int | bool | list | dict | None = None


class LSPResponseMessage(LSPMessage):
    """
    LSPResponseMessage
    """

    # The request id.
    id: int | str
    # The result of a request. This member is REQUIRED on success.
    # This member MUST NOT exist if there was an error invoking the method.
    result: str | int | bool | list | dict | None = None
    # The error object in case a request fails.
    error: LSPResponseError | None = None


class LSPNotificationMessage(LSPMessage):
    """
    LSPNotificationMessage
    """

    method: str
    params: list | dict | None = None


class ErrorCodes(Enum):
    # Defined by JSON-RPC
    ParseError = -32700
    InvalidRequest = -32600
    MethodNotFound = -32601
    InvalidParams = -32602
    InternalError = -32603

    # JSON-RPC reserved error codes range
    jsonrpcReservedErrorRangeStart = -32099
    serverErrorStart = jsonrpcReservedErrorRangeStart  # Deprecated, use jsonrpcReservedErrorRangeStart
    ServerNotInitialized = -32002
    UnknownErrorCode = -32001
    jsonrpcReservedErrorRangeEnd = -32000
    serverErrorEnd = jsonrpcReservedErrorRangeEnd  # Deprecated, use jsonrpcReservedErrorRangeEnd

    # LSP reserved error codes range
    lspReservedErrorRangeStart = -32899
    RequestFailed = -32803
    ServerCancelled = -32802
    ContentModified = -32801
    RequestCancelled = -32800
    lspReservedErrorRangeEnd = -32800
