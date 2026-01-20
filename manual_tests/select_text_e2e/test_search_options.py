"""
Search Options Tests

测试 word:select:text 的搜索选项功能。

Usage:
    uv run python manual_tests/select_text_e2e/test_search_options.py --test 1
    uv run python manual_tests/select_text_e2e/test_search_options.py --test all
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
    search_options: dict | None = None,
    selection_mode: str = "select",
    select_index: int = 1,
    wait_seconds: int = 3,
) -> bool:
    """执行文本选择动作

    Args:
        workspace: Office Workspace 实例
        document_uri: 目标文档 URI
        search_text: 要搜索的文本
        search_options: 搜索选项（matchCase, matchWholeWord, matchWildcards）
        selection_mode: 选择模式（select/start/end）
        select_index: 要选择的匹配索引（1-based）
        wait_seconds: 等待秒数

    Returns:
        bool: 操作是否成功

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

    return result.success


async def test_1_match_case_true() -> None:
    """测试 1: 区分大小写"""
    print("\n" + "=" * 60)
    print("测试 1: 区分大小写")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入 'Hello', 'HELLO', 'hello'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 只匹配大写 HELLO
        await select_text(
            workspace,
            document_uri,
            search_text="HELLO",
            search_options={"matchCase": True},
        )

        print("\n✅ 测试完成: 应该只选中大写的 'HELLO'")


async def test_2_match_case_false() -> None:
    """测试 2: 不区分大小写"""
    print("\n" + "=" * 60)
    print("测试 2: 不区分大小写")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入 'Test', 'TEST', 'test'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        # 匹配所有大小写组合
        await select_text(
            workspace,
            document_uri,
            search_text="test",
            search_options={"matchCase": False},
        )

        print("\n✅ 测试完成: 应该匹配所有大小写组合的 'test'")


async def test_3_match_whole_word() -> None:
    """测试 3: 全字匹配"""
    print("\n" + "=" * 60)
    print("测试 3: 全字匹配")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入 'test', 'test123', 'mytest', 'testing'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        # 只匹配完整的 test 单词
        await select_text(
            workspace,
            document_uri,
            search_text="test",
            search_options={"matchWholeWord": True},
        )

        print("\n✅ 测试完成: 应该只选中完整的 'test' 单词")


async def test_4_match_wildcards() -> None:
    """测试 4: 通配符搜索"""
    print("\n" + "=" * 60)
    print("测试 4: 通配符搜索")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入 'test1', 'test2', 'test123', 'mytest'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        # 使用通配符匹配 test 开头的单词
        await select_text(
            workspace,
            document_uri,
            search_text="test*",
            search_options={"matchWildcards": True},
        )

        print("\n✅ 测试完成: 应该选中所有 'test' 开头的单词")


async def test_5_combined_options() -> None:
    """测试 5: 组合搜索选项"""
    print("\n" + "=" * 60)
    print("测试 5: 组合搜索选项")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入:")
    print("   - 'Pattern' (独立单词)")
    print("   - 'pattern123' (不完整)")
    print("   - 'PATTERN' (大写)")
    print("   - 'myPattern' (后缀)")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        # 组合: 全字匹配 + 区分大小写
        await select_text(
            workspace,
            document_uri,
            search_text="Pattern",
            search_options={
                "matchCase": True,
                "matchWholeWord": True,
                "matchWildcards": False,
            },
        )

        print("\n✅ 测试完成: 应该只选中完整的 'Pattern' (区分大小写)")


async def main() -> None:
    if len(sys.argv) < 2 or sys.argv[1] != "--test":
        print("Usage: python test_search_options.py --test <1-5|all>")
        return

    test_arg = sys.argv[2] if len(sys.argv) > 2 else "1"

    tests = {
        "1": test_1_match_case_true,
        "2": test_2_match_case_false,
        "3": test_3_match_whole_word,
        "4": test_4_match_wildcards,
        "5": test_5_combined_options,
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
        print("   可用测试: 1-5, all")


if __name__ == "__main__":
    asyncio.run(main())
