"""office://workspace/documents Resource"""

from __future__ import annotations

import json

from office4ai.a2c_smcp.resources.base import BaseResource
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


class ConnectedDocumentsResource(BaseResource):
    """
    已连接文档列表资源 | Connected Documents Resource

    返回当前通过 Socket.IO 连接到 Workspace 的所有文档 URI。
    AI Agent 可通过此资源发现可操作的文档。
    """

    def __init__(self, workspace: OfficeWorkspace) -> None:
        self.workspace = workspace

    @property
    def uri(self) -> str:
        return "office://workspace/documents"

    @property
    def base_uri(self) -> str:
        return "office://workspace/documents"

    @property
    def name(self) -> str:
        return "Connected Documents"

    @property
    def description(self) -> str:
        return (
            "List of document URIs currently connected to the workspace via Office Add-In. "
            "Read this resource to discover which documents are available for tool operations."
        )

    @property
    def mime_type(self) -> str:
        return "application/json"

    async def read(self) -> str:
        documents = self.workspace.get_connected_documents()
        return json.dumps({"documents": documents, "count": len(documents)}, ensure_ascii=False)
