"""
Test Helper Functions

共享的测试辅助函数，确保所有测试文件使用一致的逻辑。
"""

import asyncio
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from typing import Any

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


def _log(emoji: str, message: str) -> None:
    """统一的日志格式"""
    print(f"\n{emoji} {message}")


async def _check_connection(workspace: OfficeWorkspace, timeout: float = 30.0) -> bool:
    """检查并等待连接（内部使用）"""
    _log("⏳", "等待 Word Add-In 连接...")
    if not await workspace.wait_for_addin_connection(timeout=timeout):
        _log("❌", "超时：未检测到 Add-In 连接")
        return False
    return True


def _get_first_document(workspace: OfficeWorkspace) -> str | None:
    """获取第一个已连接文档（内部使用）"""
    documents = workspace.get_connected_documents()
    if not documents:
        _log("❌", "未找到已连接文档")
        return None
    return documents[0]


@asynccontextmanager
async def workspace_context(host: str = "127.0.0.1", port: int = 3000) -> AsyncIterator[OfficeWorkspace]:
    """Workspace 上下文管理器（基础版本）"""
    workspace = OfficeWorkspace(host=host, port=port)
    try:
        await workspace.start()
        yield workspace
    finally:
        await workspace.stop()


@asynccontextmanager
async def ready_workspace(
    host: str = "127.0.0.1",
    port: int = 3000,
    timeout: float = 30.0,
) -> AsyncIterator[tuple[OfficeWorkspace, str]]:
    """
    完整的 Workspace 上下文管理器：启动 → 等待连接 → 获取文档 URI

    Yields:
        (workspace, document_uri): 已连接的 workspace 和第一个文档 URI

    Raises:
        RuntimeError: 如果连接失败或未找到文档
    """
    async with workspace_context(host, port) as workspace:
        if not await _check_connection(workspace, timeout):
            raise RuntimeError("Add-In 连接失败")

        doc_uri = _get_first_document(workspace)
        if not doc_uri:
            raise RuntimeError("未找到已连接文档")

        _log("✅", f"已连接文档: {doc_uri}")
        yield workspace, doc_uri


async def select_text(
    workspace: OfficeWorkspace,
    document_uri: str,
    search_text: str,
    selection_mode: str = "select",
    select_index: int = 1,
    search_options: dict | None = None,
    wait_seconds: int = 3,
) -> tuple[bool, dict | None, str | None]:
    """
    执行文本选择动作

    Args:
        workspace: Office Workspace 实例
        document_uri: 目标文档 URI
        search_text: 要搜索的文本
        selection_mode: 选择模式（select/start/end）
        select_index: 要选择的匹配索引（1-based）
        search_options: 搜索选项（matchCase, matchWholeWord, matchWildcards）
        wait_seconds: 等待秒数

    Returns:
        (success, data, error): 操作是否成功、返回数据和错误消息
    """
    print(f"\n📝 搜索文本: '{search_text[:50]}{'...' if len(search_text) > 50 else ''}'")
    print(f"📝 选择模式: {selection_mode}")
    print(f"📝 选择索引: {select_index}")
    if search_options:
        print(f"📝 搜索选项: {search_options}")

    # ✅ 使用 snake_case（符合 Python 约定），DTO 系统会自动转换为 camelCase
    params: dict[str, Any] = {
        "document_uri": document_uri,
        "search_text": search_text,
        "selection_mode": selection_mode,
        "select_index": select_index,
    }

    if search_options:
        params["search_options"] = search_options

    action = OfficeAction(
        category="word",
        action_name="select:text",
        params=params,
    )

    result = await workspace.execute(action)

    if result.success:
        print(f"✅ 选择成功")
        print(f"   返回数据: {result.data}")
        if result.data:
            if "matchCount" in result.data:
                print(f"   匹配数量: {result.data['matchCount']}")
            if "selectedIndex" in result.data:
                print(f"   选中索引: {result.data['selectedIndex']}")
            if "selectedText" in result.data:
                print(f"   选中文本: '{result.data['selectedText']}'")
    else:
        print(f"❌ 选择失败: {result.error}")

    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)

    return result.success, result.data, result.error
