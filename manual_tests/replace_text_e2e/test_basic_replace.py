"""
Basic Text Replace Test

测试基本的文本查找和替换功能。

测试场景:
1. 简单文本替换（全部）
2. 简单文本替换（首个）
3. 替换为空（删除）
4. 多行文本替换
5. 特殊字符替换
6. 长文本替换

Usage:
    # 运行单个测试
    uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test 1

    # 运行全部测试
    uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test all
"""

import asyncio
import sys
from contextlib import asynccontextmanager
from typing import Any, AsyncIterator

# Add project root to path
sys.path.insert(0, "/Users/jqq/PycharmProjects/office4ai")

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


# ==============================================================================
# 辅助函数和上下文管理器
# ==============================================================================


@asynccontextmanager
async def workspace_context(host: str = "127.0.0.1", port: int = 3000) -> AsyncIterator[OfficeWorkspace]:
    """
    Workspace 上下文管理器，自动处理启动和停止

    Args:
        host: WebSocket 服务器地址
        port: WebSocket 服务器端口

    Yields:
        OfficeWorkspace: 已启动并连接的 workspace 实例
    """
    workspace = OfficeWorkspace(host=host, port=port)
    try:
        await workspace.start()
        yield workspace
    finally:
        await workspace.stop()


async def wait_for_connection(workspace: OfficeWorkspace, timeout: float = 30.0) -> bool:
    """
    等待 Add-In 连接

    Args:
        workspace: Workspace 实例
        timeout: 超时时间（秒）

    Returns:
        bool: 是否成功连接
    """
    print("\n⏳ 等待 Word Add-In 连接...")
    connected = await workspace.wait_for_addin_connection(timeout=timeout)
    if not connected:
        print("❌ 超时：未检测到 Add-In 连接")
        return False
    return True


def get_document_uri(workspace: OfficeWorkspace) -> str | None:
    """
    获取已连接文档的 URI

    Args:
        workspace: Workspace 实例

    Returns:
        Optional[str]: 文档 URI，如果未找到则返回 None
    """
    documents = workspace.get_connected_documents()
    if not documents:
        print("❌ 未找到已连接文档")
        return None
    return documents[0]


async def replace_text(
    workspace: OfficeWorkspace,
    document_uri: str,
    search_text: str,
    replace_text: str,
    options: dict | None = None,
    wait_seconds: int = 3,
) -> bool:
    """
    执行文本替换动作

    Args:
        workspace: Workspace 实例
        document_uri: 文档 URI
        search_text: 要搜索的文本
        replace_text: 替换文本
        options: 替换选项
        wait_seconds: 等待时间（秒）

    Returns:
        bool: 是否成功
    """
    print(f"\n📝 查找文本: '{search_text[:50]}{'...' if len(search_text) > 50 else ''}'")
    print(f"📝 替换文本: '{replace_text[:50]}{'...' if len(replace_text) > 50 else ''}'")
    if options:
        print(f"📝 选项: {options}")

    params: dict[str, Any] = {
        "document_uri": document_uri,
        "searchText": search_text,
        "replaceText": replace_text,
    }

    if options:
        params["options"] = options

    action = OfficeAction(
        category="word",
        action_name="replace:text",
        params=params,
    )

    result = await workspace.execute(action)

    if result.success:
        print(f"✅ 替换成功")
        print(f"   返回数据: {result.data}")
        if result.data and "replaceCount" in result.data:
            print(f"   替换次数: {result.data['replaceCount']}")
    else:
        print(f"❌ 替换失败: {result.error}")
        if result.error:
            print(f"   错误码: {result.error}")

    # 等待一段时间让用户观察结果
    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)

    return result.success


# ==============================================================================
# 测试场景
# ==============================================================================


async def test_1_replace_all() -> None:
    """测试 1: 简单文本替换（全部）"""
    print("\n" + "=" * 60)
    print("测试 1: 简单文本替换（全部）")
    print("=" * 60)
    print("\n📋 测试说明:")
    print("   - 搜索文档中所有的 'old'")
    print("   - 替换为 'new'")
    print("   - 预期: 所有匹配项都被替换")
    print("\n📋 准备工作:")
    print("   1. 在 Word 文档中输入多次 'old' (如 5 次)")
    print("   2. 保存文档")
    print("   3. 运行测试")

    async with workspace_context() as workspace:
        # 等待连接
        if not await wait_for_connection(workspace):
            return

        # 获取文档 URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 执行替换
        await replace_text(
            workspace,
            document_uri,
            search_text="old",
            replace_text="new",
            options={"replaceAll": True},
            wait_seconds=3,
        )

        print("\n✅ 测试完成")
        print("   请在 Word 中检查: 所有的 'old' 都应该被替换为 'new'")


async def test_2_replace_first() -> None:
    """测试 2: 简单文本替换（首个）"""
    print("\n" + "=" * 60)
    print("测试 2: 简单文本替换（首个）")
    print("=" * 60)
    print("\n📋 测试说明:")
    print("   - 搜索文档中的 'test'")
    print("   - 仅替换第一个为 'exam'")
    print("   - 预期: 只有第一个匹配项被替换")
    print("\n📋 准备工作:")
    print("   1. 在 Word 文档中输入多次 'test' (如 5 次)")
    print("   2. 保存文档")
    print("   3. 运行测试")

    async with workspace_context() as workspace:
        # 等待连接
        if not await wait_for_connection(workspace):
            return

        # 获取文档 URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 执行替换
        await replace_text(
            workspace,
            document_uri,
            search_text="test",
            replace_text="exam",
            options={"replaceAll": False},
            wait_seconds=3,
        )

        print("\n✅ 测试完成")
        print("   请在 Word 中检查: 只有第一个 'test' 被替换为 'exam'")


async def test_3_replace_with_empty() -> None:
    """测试 3: 替换为空（删除）"""
    print("\n" + "=" * 60)
    print("测试 3: 替换为空（删除）")
    print("=" * 60)
    print("\n📋 测试说明:")
    print("   - 搜索文档中的 'delete'")
    print("   - 替换为空字符串（删除匹配的文本）")
    print("   - 预期: 所有 'delete' 被删除")
    print("\n📋 准备工作:")
    print("   1. 在 Word 文档中输入多次 'delete' (如 3 次)")
    print("   2. 保存文档")
    print("   3. 运行测试")

    async with workspace_context() as workspace:
        # 等待连接
        if not await wait_for_connection(workspace):
            return

        # 获取文档 URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 执行替换
        await replace_text(
            workspace,
            document_uri,
            search_text="delete",
            replace_text="",
            options={"replaceAll": True},
            wait_seconds=3,
        )

        print("\n✅ 测试完成")
        print("   请在 Word 中检查: 所有的 'delete' 都应该被删除")


async def test_4_multiline_replace() -> None:
    """测试 4: 多行文本替换"""
    print("\n" + "=" * 60)
    print("测试 4: 多行文本替换")
    print("=" * 60)
    print("\n📋 测试说明:")
    print("   - 搜索多行文本 'line1\\nline2'")
    print("   - 替换为 'new\\ncontent'")
    print("   - 预期: 多行文本被正确替换")
    print("\n📋 准备工作:")
    print("   1. 在 Word 文档中输入:")
    print("      line1")
    print("      line2")
    print("   2. 重复输入 3 次")
    print("   3. 保存文档")
    print("   4. 运行测试")

    async with workspace_context() as workspace:
        # 等待连接
        if not await wait_for_connection(workspace):
            return

        # 获取文档 URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 执行替换
        await replace_text(
            workspace,
            document_uri,
            search_text="line1\nline2",
            replace_text="new\ncontent",
            options={"replaceAll": True},
            wait_seconds=3,
        )

        print("\n✅ 测试完成")
        print("   请在 Word 中检查: 所有 'line1\\nline2' 都应该被替换为 'new\\ncontent'")


async def test_5_special_characters() -> None:
    """测试 5: 特殊字符替换"""
    print("\n" + "=" * 60)
    print("测试 5: 特殊字符替换")
    print("=" * 60)
    print("\n📋 测试说明:")
    print("   - 搜索包含特殊字符的文本 'Café'")
    print("   - 替换为 'Coffee'")
    print("   - 预期: 特殊字符被正确处理")
    print("\n📋 准备工作:")
    print("   1. 在 Word 文档中输入多次 'Café' (如 3 次)")
    print("   2. 保存文档")
    print("   3. 运行测试")

    async with workspace_context() as workspace:
        # 等待连接
        if not await wait_for_connection(workspace):
            return

        # 获取文档 URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 执行替换
        await replace_text(
            workspace,
            document_uri,
            search_text="Café",
            replace_text="Coffee",
            options={"replaceAll": True},
            wait_seconds=3,
        )

        print("\n✅ 测试完成")
        print("   请在 Word 中检查: 所有的 'Café' 都应该被替换为 'Coffee'")


async def test_6_long_text_replace() -> None:
    """测试 6: 长文本替换"""
    print("\n" + "=" * 60)
    print("测试 6: 长文本替换")
    print("=" * 60)
    print("\n📋 测试说明:")
    print("   - 搜索较长的文本段落")
    print("   - 替换为另一个长文本段落")
    print("   - 预期: 长文本被正确替换")
    print("\n📋 准备工作:")
    print("   1. 在 Word 文档中输入以下文本 2 次:")
    print("      'This is a long paragraph of text that should be replaced with another long paragraph. It contains multiple sentences and various punctuation marks.'")
    print("   2. 保存文档")
    print("   3. 运行测试")

    async with workspace_context() as workspace:
        # 等待连接
        if not await wait_for_connection(workspace):
            return

        # 获取文档 URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 准备长文本
        search_text = "This is a long paragraph of text that should be replaced with another long paragraph. It contains multiple sentences and various punctuation marks."
        replacement_text = "Here is another lengthy paragraph that serves as the replacement text. It also has multiple sentences and demonstrates the replace functionality."

        # 执行替换
        await replace_text(
            workspace,
            document_uri,
            search_text=search_text,
            replace_text=replacement_text,
            options={"replaceAll": True},
            wait_seconds=3,
        )

        print("\n✅ 测试完成")
        print("   请在 Word 中检查: 长文本应该被正确替换")


# ==============================================================================
# 主函数
# ==============================================================================


async def main() -> None:
    """主函数"""
    if len(sys.argv) < 2 or sys.argv[1] != "--test":
        print("Usage: python test_basic_replace.py --test <1-6|all>")
        return

    test_arg = sys.argv[2] if len(sys.argv) > 2 else "1"

    tests = {
        "1": test_1_replace_all,
        "2": test_2_replace_first,
        "3": test_3_replace_with_empty,
        "4": test_4_multiline_replace,
        "5": test_5_special_characters,
        "6": test_6_long_text_replace,
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
        print("   可用测试: 1-6, all")


if __name__ == "__main__":
    asyncio.run(main())
