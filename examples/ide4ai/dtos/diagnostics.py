# filename: diagnostics.py
# @Time    : 2025/10/27 14:38
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
LSP 诊断相关的数据传输对象 / LSP Diagnostics Data Transfer Objects

定义了 Document Diagnostics 和 Workspace Diagnostics 相关的 Pydantic 模型。
Defines Pydantic models for Document Diagnostics and Workspace Diagnostics.
"""

from typing import Literal

from pydantic import BaseModel, Field


class DiagnosticSeverity:
    """
    诊断严重程度 / Diagnostic severity levels
    """

    Error = 1
    Warning = 2
    Information = 3
    Hint = 4


class DiagnosticTag:
    """
    诊断标签 / Diagnostic tags
    """

    Unnecessary = 1  # 未使用或不必要的代码 / Unused or unnecessary code
    Deprecated = 2  # 已弃用的代码 / Deprecated code


class Position(BaseModel):
    """
    文档中的位置 / Position in a document
    """

    line: int = Field(..., description="行号（0-based）/ Line number (0-based)")
    character: int = Field(..., description="字符偏移（0-based）/ Character offset (0-based)")


class Range(BaseModel):
    """
    文档中的范围 / Range in a document
    """

    start: Position = Field(..., description="范围起始位置 / Start position")
    end: Position = Field(..., description="范围结束位置 / End position")


class DiagnosticRelatedInformation(BaseModel):
    """
    诊断相关信息 / Diagnostic related information
    """

    location: dict = Field(..., description="相关位置 / Related location")
    message: str = Field(..., description="相关信息消息 / Related information message")


class CodeDescription(BaseModel):
    """
    代码描述 / Code description
    """

    href: str = Field(..., description="指向代码描述的URI / URI pointing to code description")


class Diagnostic(BaseModel):
    """
    诊断信息 / Diagnostic information
    """

    range: Range = Field(..., description="诊断适用的范围 / Range to which the diagnostic applies")
    severity: int | None = Field(None, description="诊断严重程度 / Diagnostic severity")
    code: int | str | None = Field(None, description="诊断代码 / Diagnostic code")
    codeDescription: CodeDescription | None = Field(None, description="代码描述 / Code description")
    source: str | None = Field(None, description="诊断来源 / Diagnostic source")
    message: str = Field(..., description="诊断消息 / Diagnostic message")
    tags: list[int] | None = Field(None, description="诊断标签 / Diagnostic tags")
    relatedInformation: list[DiagnosticRelatedInformation] | None = Field(
        None,
        description="相关诊断信息 / Related diagnostic information",
    )
    data: dict | None = Field(None, description="额外数据 / Additional data")


class DocumentDiagnosticReportKind:
    """
    文档诊断报告类型 / Document diagnostic report kinds
    """

    Full = "full"  # 完整的诊断报告 / Full diagnostic report
    Unchanged = "unchanged"  # 未改变的诊断报告 / Unchanged diagnostic report


class FullDocumentDiagnosticReport(BaseModel):
    """
    完整的文档诊断报告 / Full document diagnostic report
    """

    kind: Literal["full"] = Field(
        "full",
        description="报告类型：完整 / Report kind: full",
    )
    resultId: str | None = Field(
        None,
        description="结果ID，用于后续请求 / Result ID for subsequent requests",
    )
    items: list[Diagnostic] = Field(..., description="诊断项列表 / List of diagnostic items")


class UnchangedDocumentDiagnosticReport(BaseModel):
    """
    未改变的文档诊断报告 / Unchanged document diagnostic report
    """

    kind: Literal["unchanged"] = Field(
        "unchanged",
        description="报告类型：未改变 / Report kind: unchanged",
    )
    resultId: str = Field(..., description="结果ID / Result ID")


class RelatedFullDocumentDiagnosticReport(FullDocumentDiagnosticReport):
    """
    带相关文档的完整诊断报告 / Full diagnostic report with related documents
    """

    relatedDocuments: dict[str, FullDocumentDiagnosticReport | UnchangedDocumentDiagnosticReport] | None = Field(
        None,
        description="相关文档的诊断信息 / Diagnostics of related documents",
    )


class RelatedUnchangedDocumentDiagnosticReport(UnchangedDocumentDiagnosticReport):
    """
    带相关文档的未改变诊断报告 / Unchanged diagnostic report with related documents
    """

    relatedDocuments: dict[str, FullDocumentDiagnosticReport | UnchangedDocumentDiagnosticReport] | None = Field(
        None,
        description="相关文档的诊断信息 / Diagnostics of related documents",
    )


# 文档诊断报告类型联合 / Document diagnostic report type union
DocumentDiagnosticReport = RelatedFullDocumentDiagnosticReport | RelatedUnchangedDocumentDiagnosticReport


class DocumentDiagnosticParams(BaseModel):
    """
    文档诊断请求参数 / Document diagnostic request parameters
    """

    textDocument: dict = Field(..., description="文本文档标识 / Text document identifier")
    identifier: str | None = Field(
        default=None, description="注册时提供的额外标识符 / Additional identifier from registration"
    )
    previousResultId: str | None = Field(
        default=None,
        description="上一次响应的结果ID / Result ID from previous response",
    )


class WorkspaceFullDocumentDiagnosticReport(FullDocumentDiagnosticReport):
    """
    工作区完整文档诊断报告 / Workspace full document diagnostic report
    """

    uri: str = Field(..., description="文档URI / Document URI")
    version: int | None = Field(default=None, description="文档版本号 / Document version number")


class WorkspaceUnchangedDocumentDiagnosticReport(UnchangedDocumentDiagnosticReport):
    """
    工作区未改变文档诊断报告 / Workspace unchanged document diagnostic report
    """

    uri: str = Field(..., description="文档URI / Document URI")
    version: int | None = Field(default=None, description="文档版本号 / Document version number")


# 工作区文档诊断报告类型联合 / Workspace document diagnostic report type union
WorkspaceDocumentDiagnosticReport = WorkspaceFullDocumentDiagnosticReport | WorkspaceUnchangedDocumentDiagnosticReport


class WorkspaceDiagnosticReport(BaseModel):
    """
    工作区诊断报告 / Workspace diagnostic report
    """

    items: list[WorkspaceDocumentDiagnosticReport] = Field(
        ...,
        description="工作区文档诊断报告列表 / List of workspace document diagnostic reports",
    )


class PreviousResultId(BaseModel):
    """
    上一次结果ID / Previous result ID
    """

    uri: str = Field(..., description="文档URI / Document URI")
    value: str = Field(..., description="上一次结果ID的值 / Value of previous result ID")


class WorkspaceDiagnosticParams(BaseModel):
    """
    工作区诊断请求参数 / Workspace diagnostic request parameters
    """

    identifier: str | None = Field(
        default=None, description="注册时提供的额外标识符 / Additional identifier from registration"
    )
    previousResultIds: list[PreviousResultId] = Field(
        default_factory=list,
        description="已知的诊断报告及其结果ID / Known diagnostic reports with result IDs",
    )


class DiagnosticServerCancellationData(BaseModel):
    """
    诊断服务器取消数据 / Diagnostic server cancellation data
    """

    retriggerRequest: bool = Field(..., description="是否应重新触发请求 / Whether to retrigger the request")


class DocumentDiagnosticReportPartialResult(BaseModel):
    """
    文档诊断报告部分结果 / Document diagnostic report partial result
    """

    relatedDocuments: dict[str, FullDocumentDiagnosticReport | UnchangedDocumentDiagnosticReport] = Field(
        ...,
        description="相关文档的诊断 / Diagnostics of related documents",
    )


class WorkspaceDiagnosticReportPartialResult(BaseModel):
    """
    工作区诊断报告部分结果 / Workspace diagnostic report partial result
    """

    items: list[WorkspaceDocumentDiagnosticReport] = Field(
        ...,
        description="工作区文档诊断报告列表 / List of workspace document diagnostic reports",
    )
