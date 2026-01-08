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

# Excel DTOs
from office4ai.environment.workspace.dtos.excel import (
    ExcelGetRangeRequest,
    ExcelGetSelectedRangeRequest,
    ExcelGetUsedRangeRequest,
    ExcelInsertChartRequest,
    ExcelInsertTableRequest,
    ExcelSetCellValueRequest,
    ExcelSetRangeRequest,
)

# PPT DTOs
from office4ai.environment.workspace.dtos.ppt import (
    PptDeleteSlideRequest,
    PptGetCurrentSlideElementsRequest,
    PptGetSlideElementsRequest,
    PptGetSlideScreenshotRequest,
    PptInsertImageRequest,
    PptInsertShapeRequest,
    PptInsertTableRequest,
    PptInsertTextRequest,
    PptMoveSlideRequest,
    PptUpdateTextBoxRequest,
)

# Word DTOs
from office4ai.environment.workspace.dtos.word import (
    WordAppendTextRequest,
    WordExportContentRequest,
    WordGetDocumentStatsRequest,
    WordGetDocumentStructureRequest,
    WordGetSelectedContentRequest,
    WordGetVisibleContentRequest,
    WordInsertEquationRequest,
    WordInsertImageRequest,
    WordInsertTableRequest,
    WordInsertTextRequest,
    WordInsertTOCRequest,
    WordReplaceSelectionRequest,
    WordReplaceTextRequest,
)

__all__ = [
    # Common
    "BaseRequest",
    "BaseResponse",
    "ErrorResponse",
    "ErrorCode",
    # Word
    "WordGetSelectedContentRequest",
    "WordGetVisibleContentRequest",
    "WordGetDocumentStructureRequest",
    "WordGetDocumentStatsRequest",
    "WordInsertTextRequest",
    "WordReplaceSelectionRequest",
    "WordReplaceTextRequest",
    "WordAppendTextRequest",
    "WordInsertImageRequest",
    "WordInsertTableRequest",
    "WordInsertEquationRequest",
    "WordInsertTOCRequest",
    "WordExportContentRequest",
    # Excel
    "ExcelGetSelectedRangeRequest",
    "ExcelGetUsedRangeRequest",
    "ExcelSetCellValueRequest",
    "ExcelInsertTableRequest",
    "ExcelGetRangeRequest",
    "ExcelSetRangeRequest",
    "ExcelInsertChartRequest",
    # PPT
    "PptGetCurrentSlideElementsRequest",
    "PptGetSlideElementsRequest",
    "PptGetSlideScreenshotRequest",
    "PptInsertTextRequest",
    "PptInsertImageRequest",
    "PptInsertTableRequest",
    "PptInsertShapeRequest",
    "PptDeleteSlideRequest",
    "PptMoveSlideRequest",
    "PptUpdateTextBoxRequest",
]
