"""
Socket.IO Request Wrapper

Automatically wraps business parameters into BaseRequest-compatible DTOs
with auto-generated request IDs and timestamps.
"""

import uuid
from typing import Any

from pydantic import ValidationError

from office4ai.environment.workspace.dtos import excel, ppt, word
from office4ai.environment.workspace.dtos.common import BaseRequest

# Event name → DTO class mapping
REQUEST_DTO_REGISTRY: dict[str, type[BaseRequest]] = {
    # Word events (13)
    "word:get:selectedContent": word.WordGetSelectedContentRequest,
    "word:get:visibleContent": word.WordGetVisibleContentRequest,
    "word:get:documentStructure": word.WordGetDocumentStructureRequest,
    "word:get:documentStats": word.WordGetDocumentStatsRequest,
    "word:insert:text": word.WordInsertTextRequest,
    "word:replace:selection": word.WordReplaceSelectionRequest,
    "word:replace:text": word.WordReplaceTextRequest,
    "word:append:text": word.WordAppendTextRequest,
    "word:insert:image": word.WordInsertImageRequest,
    "word:insert:table": word.WordInsertTableRequest,
    "word:insert:equation": word.WordInsertEquationRequest,
    "word:insert:toc": word.WordInsertTOCRequest,
    "word:export:content": word.WordExportContentRequest,
    # Excel events (10)
    "excel:get:selectedRange": excel.ExcelGetSelectedRangeRequest,
    "excel:get:usedRange": excel.ExcelGetUsedRangeRequest,
    "excel:set:cellValue": excel.ExcelSetCellValueRequest,
    "excel:insert:table": excel.ExcelInsertTableRequest,
    "excel:get:range": excel.ExcelGetRangeRequest,
    "excel:set:range": excel.ExcelSetRangeRequest,
    "excel:insert:chart": excel.ExcelInsertChartRequest,
    # PPT events (10)
    "ppt:get:currentSlideElements": ppt.PptGetCurrentSlideElementsRequest,
    "ppt:get:slideElements": ppt.PptGetSlideElementsRequest,
    "ppt:get:slideScreenshot": ppt.PptGetSlideScreenshotRequest,
    "ppt:insert:text": ppt.PptInsertTextRequest,
    "ppt:insert:image": ppt.PptInsertImageRequest,
    "ppt:insert:table": ppt.PptInsertTableRequest,
    "ppt:insert:shape": ppt.PptInsertShapeRequest,
    "ppt:delete:slide": ppt.PptDeleteSlideRequest,
    "ppt:move:slide": ppt.PptMoveSlideRequest,
    "ppt:update:textBox": ppt.PptUpdateTextBoxRequest,
}


class RequestWrapperError(Exception):
    """Raised when request wrapping fails"""

    pass


def wrap_request(
    event: str,
    business_params: dict[str, Any],
    document_uri: str | None = None,
) -> dict[str, Any]:
    """
    Wrap business parameters into a BaseRequest-compatible DTO.

    This function:
    1. Looks up the appropriate DTO class for the event
    2. Auto-generates a unique requestId (UUID4)
    3. Extracts document_uri from business_params if not provided
    4. Merges business params with BaseRequest fields
    5. Validates using Pydantic
    6. Returns camelCase JSON for Socket.IO transmission

    Args:
        event: Event name (e.g., "word:get:selectedContent")
        business_params: Business-specific parameters
        document_uri: Optional document URI (extracted from params if not provided)

    Returns:
        dict: JSON-ready dict with camelCase keys

    Raises:
        RequestWrapperError: If event not registered or validation fails

    Example:
        >>> params = {"options": {"includeText": True}}
        >>> wrapped = wrap_request("word:get:selectedContent", params, "file:///test.docx")
        >>> print(wrapped)
        {
            "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
            "documentUri": "file:///test.docx",
            "timestamp": 1234567890000,
            "options": {"includeText": True}
        }
    """
    # Step 1: Look up DTO class
    dto_class = REQUEST_DTO_REGISTRY.get(event)
    if not dto_class:
        raise RequestWrapperError(f"Unknown event '{event}'. Not registered in REQUEST_DTO_REGISTRY.")

    # Step 2: Extract document_uri
    if not document_uri:
        document_uri = business_params.get("document_uri")
        if not document_uri:
            raise RequestWrapperError("document_uri must be provided either as argument or in business_params")

    # Step 3: Generate unique request ID
    request_id = str(uuid.uuid4())

    # Step 4: Prepare complete parameters
    complete_params = {
        "request_id": request_id,  # snake_case for internal Python
        "document_uri": document_uri,
        **business_params,  # Business parameters (may include document_uri again, that's fine)
    }

    # Step 5: Validate using Pydantic DTO
    try:
        # Create DTO instance (validates all fields)
        dto_instance = dto_class(**complete_params)
    except ValidationError as e:
        raise RequestWrapperError(f"Failed to validate request for event '{event}': {e}") from e
    except Exception as e:
        raise RequestWrapperError(f"Unexpected error creating DTO for event '{event}': {e}") from e

    # Step 6: Convert to camelCase JSON (Pydantic handles alias conversion)
    return dto_instance.model_dump(by_alias=True, exclude_none=True)


def is_wrappable_event(event: str) -> bool:
    """Check if an event has a registered DTO wrapper"""
    return event in REQUEST_DTO_REGISTRY


def get_registered_events() -> list[str]:
    """Get list of all registered event names"""
    return sorted(REQUEST_DTO_REGISTRY.keys())
