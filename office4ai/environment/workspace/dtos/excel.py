"""
Excel Socket.IO DTOs

Defines data structures for Excel-specific Socket.IO events.
"""

from typing import Any, ClassVar, Literal, Optional

from pydantic import BaseModel, Field

from .common import BaseRequest

# ============================================================================
# Range Operations DTOs
# ============================================================================


class ExcelGetSelectedRangeRequest(BaseRequest):
    """
    Request to get selected range from Excel worksheet.
    """

    event_name: ClassVar[str] = "excel:get:selectedRange"

    pass


class ExcelGetUsedRangeRequest(BaseRequest):
    """
    Request to get used range from Excel worksheet.
    """

    event_name: ClassVar[str] = "excel:get:usedRange"

    worksheetName: str | None = Field(default=None, description="Worksheet name (default: active)")


class ExcelSetCellValueRequest(BaseRequest):
    """
    Request to set cell value in Excel worksheet.
    """

    event_name: ClassVar[str] = "excel:set:cellValue"

    address: str = Field(..., description="Cell address (e.g., 'A1' or 'Sheet1!A1')")
    value: Any = Field(..., description="Cell value (string, number, boolean, etc.)")
    worksheetName: str | None = Field(default=None, description="Worksheet name (default: active)")


class ExcelInsertTableRequest(BaseRequest):
    """
    Request to insert table into Excel worksheet.
    """

    event_name: ClassVar[str] = "excel:insert:table"

    options: "ExcelTableInsertOptions" = Field(..., description="Table insertion options")


class ExcelTableInsertOptions(BaseModel):
    """
    Table insertion options for Excel.
    """

    rows: int = Field(..., description="Number of rows", ge=1)
    columns: int = Field(..., description="Number of columns", ge=1)
    data: list[list[Any]] | None = Field(default=None, description="Table data")
    address: str | None = Field(default=None, description="Target address (e.g., 'A1')")
    worksheetName: str | None = Field(default=None, description="Worksheet name (default: active)")
    hasHeaders: bool = Field(default=True, description="First row contains headers")
    styleName: str | None = Field(default=None, description="Table style name")


# ============================================================================
# Additional Excel DTOs (for future expansion)
# ============================================================================


class ExcelGetRangeRequest(BaseRequest):
    """
    Request to get a specific range from Excel worksheet.
    """

    event_name: ClassVar[str] = "excel:get:range"

    address: str = Field(..., description="Range address (e.g., 'A1:C10')")
    worksheetName: str | None = Field(default=None, description="Worksheet name (default: active)")


class ExcelSetRangeRequest(BaseRequest):
    """
    Request to set values in a range.
    """

    event_name: ClassVar[str] = "excel:set:range"

    address: str = Field(..., description="Range address (e.g., 'A1:C10')")
    values: list[list[Any]] = Field(..., description="2D array of values")
    worksheetName: str | None = Field(default=None, description="Worksheet name (default: active)")


class ExcelInsertChartRequest(BaseRequest):
    """
    Request to insert chart into Excel worksheet.
    """

    event_name: ClassVar[str] = "excel:insert:chart"

    chartType: Literal[
        "Column",
        "Line",
        "Pie",
        "Bar",
        "Area",
        "Scatter",
        "Doughnut",
    ] = Field(..., description="Chart type")
    dataRange: str = Field(..., description="Data range (e.g., 'A1:C10')")
    options: Optional["ChartInsertOptions"] = Field(default=None, description="Chart options")


class ChartInsertOptions(BaseModel):
    """
    Chart insertion options.
    """

    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    width: float | None = Field(default=None, description="Width (points)")
    height: float | None = Field(default=None, description="Height (points)")
    title: str | None = Field(default=None, description="Chart title")
    worksheetName: str | None = Field(default=None, description="Worksheet name (default: active)")


# Resolve forward references
ExcelTableInsertOptions.model_rebuild()
ChartInsertOptions.model_rebuild()
