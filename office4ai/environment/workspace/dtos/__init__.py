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

# Word DTOs
from office4ai.environment.workspace.dtos.word import (
    WordGetSelectedContentRequest,
    WordGetVisibleContentRequest,
    WordGetDocumentStructureRequest,
    WordGetDocumentStatsRequest,
    WordInsertTextRequest,
    WordReplaceSelectionRequest,
    WordReplaceTextRequest,
    WordAppendTextRequest,
    WordInsertImageRequest,
    WordInsertTableRequest,
    WordInsertEquationRequest,
    WordInsertTOCRequest,
    WordExportContentRequest,
)

# Excel DTOs
from office4ai.environment.workspace.dtos.excel import (
    ExcelGetSelectedRangeRequest,
    ExcelGetUsedRangeRequest,
    ExcelSetCellValueRequest,
    ExcelInsertTableRequest,
    ExcelGetRangeRequest,
    ExcelSetRangeRequest,
    ExcelInsertChartRequest,
)

# PPT DTOs
from office4ai.environment.workspace.dtos.ppt import (
    PptGetCurrentSlideElementsRequest,
    PptGetSlideElementsRequest,
    PptGetSlideScreenshotRequest,
    PptInsertTextRequest,
    PptInsertImageRequest,
    PptInsertTableRequest,
    PptInsertShapeRequest,
    PptDeleteSlideRequest,
    PptMoveSlideRequest,
    PptUpdateTextBoxRequest,
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
