"""
Test Helper Functions

共享的测试辅助函数，确保所有测试文件使用一致的逻辑。
"""

import asyncio
import sys
from typing import Any

sys.path.insert(0, "/Users/jqq/PycharmProjects/office4ai")

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


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
