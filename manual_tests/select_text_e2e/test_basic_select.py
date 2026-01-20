"""
Basic Text Selection Test

测试基础的文本选择功能。

Usage:
    uv run python manual_tests/select_text_e2e/test_basic_select.py --test 1
    uv run python manual_tests/select_text_e2e/test_basic_select.py --test all
"""

import asyncio
import sys
from contextlib import asynccontextmanager
from typing import Any, AsyncIterator

sys.path.insert(0, "/Users/jqq/PycharmProjects/office4ai")

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@asynccontextmanager
async def workspace_context(host: str = "127.0.0.1", port: int = 3000) -> AsyncIterator[OfficeWorkspace]:
    workspace = OfficeWorkspace(host=host, port=port)
    try:
        await workspace.start()
        yield workspace
    finally:
        await workspace.stop()


async def wait_for_connection(workspace: OfficeWorkspace, timeout: float = 30.0) -> bool:
    print("\n⏳ 等待 Word Add-In 连接...")
    connected = await workspace.wait_for_addin_connection(timeout=timeout)
    if not connected:
        print("❌ 超时：未检测到 Add-In 连接")
        return False
    return True


def get_document_uri(workspace: OfficeWorkspace) -> str | None:
    documents = workspace.get_connected_documents()
    if not documents:
        print("❌ 未找到已连接文档")
        return None
    return documents[0]


async def select_text(
    workspace: OfficeWorkspace,
    document_uri: str,
    search_text: str,
    selection_mode: str = "select",
    select_index: int = 1,
    search_options: dict | None = None,
    wait_seconds: int = 3,
) -> tuple[bool, dict | None]:
    """执行文本选择动作

    Args:
        workspace: Office Workspace 实例
        document_uri: 目标文档 URI
        search_text: 要搜索的文本
        selection_mode: 选择模式（select/start/end）
        select_index: 要选择的匹配索引（1-based）
        search_options: 搜索选项（matchCase, matchWholeWord, matchWildcards）
        wait_seconds: 等待秒数

    Returns:
        (success, data): 操作是否成功和返回数据

    Note:
        参数使用 snake_case（符合 Python 约定），
        DTO 系统会自动转换为协议层的 camelCase
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

    return result.success, result.data


async def test_1_simple_selection() -> None:
    """测试 1: 简单选中文本"""
    print("\n" + "=" * 60)
    print("测试 1: 简单选中文本")
    print("=" * 60)
    print("\n📋 准备: 在 Word 文档中输入多次 'Hello World'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        success, data = await select_text(workspace, document_uri, search_text="Hello World")

        # 验证结果
        print("\n🔍 验证测试结果:")
        if success and data and data.get("matchCount", 0) > 0:
            print("   ✓ 找到并选中文本")
            print(f"\n✅ 测试通过: 'Hello World' 已被选中")
        else:
            print("   ❌ 未找到文本或操作失败")
            print("\n❌ 测试失败: 选择 'Hello World' 失败")
            print("   请确认文档中包含 'Hello World' 文本")


async def test_2_select_nth_match() -> None:
    """测试 2: 选择第N个匹配项"""
    print("\n" + "=" * 60)
    print("测试 2: 选择第N个匹配项")
    print("=" * 60)
    print("\n📋 准备: 在 Word 文档中输入 5 次 'test'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        success, data = await select_text(workspace, document_uri, search_text="test", select_index=2)

        # 验证结果
        print("\n🔍 验证测试结果:")
        if success and data:
            match_count = data.get("matchCount", 0)
            if match_count >= 2:
                print(f"   ✓ 找到 {match_count} 个匹配")
                print(f"   ✓ 成功选中第 2 个 'test'")
                print("\n✅ 测试通过: 第 2 个 'test' 已被选中")
            else:
                print(f"   ❌ 只找到 {match_count} 个 'test'，需要至少 2 个")
                print("\n❌ 测试失败: 文档中 'test' 数量不足")
        else:
            print("   ❌ 操作失败或未找到匹配")
            print("\n❌ 测试失败: 选择第 2 个 'test' 失败")
            print("   请确认文档中至少有 2 个 'test' 文本")


async def test_3_case_insensitive() -> None:
    """测试 3: 不区分大小写"""
    print("\n" + "=" * 60)
    print("测试 3: 不区分大小写")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入 'Hello', 'HELLO', 'hello'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        success, data = await select_text(
            workspace,
            document_uri,
            search_text="hello",
            search_options={"matchCase": False},
        )

        # 验证结果
        print("\n🔍 验证测试结果:")
        if success and data and data.get("matchCount", 0) > 0:
            print("   ✓ 不区分大小写匹配成功")
            print("\n✅ 测试通过: 找到大小写不同的 'hello'")
        else:
            print("   ❌ 未找到匹配")
            print("\n❌ 测试失败: 不区分大小写匹配失败")
            print("   请确认文档中包含 'Hello', 'HELLO' 或 'hello'")


async def test_4_whole_word() -> None:
    """测试 4: 全字匹配"""
    print("\n" + "=" * 60)
    print("测试 4: 全字匹配")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入 'test', 'test123', 'mytest'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        success, data = await select_text(
            workspace,
            document_uri,
            search_text="test",
            search_options={"matchWholeWord": True},
        )

        # 验证结果
        print("\n🔍 验证测试结果:")
        if success and data and data.get("matchCount", 0) > 0:
            selected_text = data.get("selectedText", "")
            if selected_text == "test":
                print("   ✓ 全字匹配成功")
                print(f"   ✓ 选中文本: '{selected_text}'")
                print("\n✅ 测试通过: 只选中完整的 'test' 单词")
            else:
                print(f"   ⚠️  选中的文本不完全是 'test': '{selected_text}'")
        else:
            print("   ❌ 未找到完整的 'test' 单词")
            print("\n❌ 测试失败: 全字匹配失败")
            print("   请确认文档中包含完整的 'test' 单词")


async def main() -> None:
    if len(sys.argv) < 2 or sys.argv[1] != "--test":
        print("Usage: python test_basic_select.py --test <1-4|all>")
        return

    test_arg = sys.argv[2] if len(sys.argv) > 2 else "1"

    tests = {
        "1": test_1_simple_selection,
        "2": test_2_select_nth_match,
        "3": test_3_case_insensitive,
        "4": test_4_whole_word,
    }

    if test_arg == "all":
        for test_num, test_func in tests.items():
            try:
                await test_func()
                print("\n" + "▓" * 60)
                print(f"✅ 测试 {test_num} 完成\n")
            except Exception as e:
                print(f"\n❌ 测试 {test_num} 失败: {e}\n")
    elif test_arg in tests:
        try:
            await tests[test_arg]()
        except Exception as e:
            print(f"\n❌ 测试失败: {e}\n")
    else:
        print(f"❌ 无效的测试编号: {test_arg}")
        print("   可用测试: 1-4, all")


if __name__ == "__main__":
    asyncio.run(main())
