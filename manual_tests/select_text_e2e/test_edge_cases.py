"""
Edge Cases Tests

测试 word:select:text 的边界情况和错误处理。

Usage:
    uv run python manual_tests/select_text_e2e/test_edge_cases.py --test 1
    uv run python manual_tests/select_text_e2e/test_edge_cases.py --test all
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
    wait_seconds: int = 3,
) -> tuple[bool, dict | None, str | None]:
    """执行文本选择动作

    Returns:
        (success, data, error): 操作是否成功、返回数据和错误消息
    """
    print(f"\n📝 搜索文本: '{search_text[:50]}{'...' if len(search_text) > 50 else ''}'")
    print(f"📝 选择模式: {selection_mode}")
    print(f"📝 选择索引: {select_index}")

    params: dict[str, Any] = {
        "document_uri": document_uri,
        "searchText": search_text,
        "selectionMode": selection_mode,
        "selectIndex": select_index,
    }

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


async def test_1_no_matches() -> None:
    """测试 1: 未找到匹配文本"""
    print("\n" + "=" * 60)
    print("测试 1: 未找到匹配文本")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入一些文本（不要包含 'NonExistentText'）")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        success, data, error = await select_text(
            workspace,
            document_uri,
            search_text="NonExistentText",
        )

        # 验证结果
        print("\n🔍 验证测试结果:")
        test_passed = True

        # 未找到文本时，协议层返回 success=False
        if not success:
            print("   ✓ 操作返回 success=False（未找到文本）")
        else:
            print("   ❌ 操作应该返回 success=False（未找到匹配）")
            test_passed = False

        # 修复后：协议层失败时，data 为空字典，错误信息在 error 字段
        if data == {}:
            print("   ✓ data 为空字典（协议层失败，无业务数据）")
        elif data is None:
            print("   ❌ 返回数据不应该为 None")
            test_passed = False
        else:
            # 如果有 data，检查 matchCount（兼容修复前的行为）
            if data.get("matchCount") == 0:
                print("   ✓ matchCount = 0（兼容修复前格式）")
            else:
                print(f"   ⚠️  matchCount = {data.get('matchCount')}（协议层失败时不应该有 data）")

        # 检查错误消息（修复后应该有清晰的错误信息）
        if error:
            print(f"   ✓ 错误消息: {error}")
            # 验证错误消息包含关键信息
            if "NonExistentText" in error or "0 matches" in error:
                print("   ✓ 错误消息包含搜索文本和匹配数")
        else:
            print("   ❌ 缺少错误消息")
            test_passed = False

        if test_passed:
            print("\n✅ 测试通过")
        else:
            print("\n❌ 测试失败")


async def test_2_empty_search_text() -> None:
    """测试 2: 空搜索文本"""
    print("\n" + "=" * 60)
    print("测试 2: 空搜索文本")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入任意文本")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        success, data, error = await select_text(
            workspace,
            document_uri,
            search_text="",  # 空字符串
        )

        # 验证结果
        print("\n🔍 验证测试结果:")
        test_passed = True

        # 空文本搜索应该失败或返回警告
        if success:
            print("   ⚠️  操作成功（空文本搜索的处理取决于实现）")
            if data and data.get("matchCount") == 0:
                print("   ✓ matchCount = 0")
        else:
            print("   ✓ 操作返回错误（符合预期）")

        print("\n✅ 测试完成")


async def test_3_out_of_bounds_index() -> None:
    """测试 3: 超出索引范围"""
    print("\n" + "=" * 60)
    print("测试 3: 超出索引范围")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入 3 次 'OutOfBounds'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 只匹配了 3 次，但请求第 10 次
        success, data, error = await select_text(
            workspace,
            document_uri,
            search_text="OutOfBounds",
            select_index=10,  # 超出范围
        )

        # 验证结果
        print("\n🔍 验证测试结果:")
        test_passed = True

        if not success:
            print("   ✓ 操作返回失败（符合预期：索引超出范围）")
        else:
            print("   ❌ 操作应该失败（索引超出范围）")
            test_passed = False

        if data:
            match_count = data.get("matchCount")
            if match_count is not None:
                if match_count == 3:
                    print(f"   ✓ matchCount = {match_count}")
                else:
                    print(f"   ⚠️  matchCount = {match_count}（预期 3）")

        if test_passed:
            print("\n✅ 测试通过")
        else:
            print("\n❌ 测试失败")


async def test_4_special_characters() -> None:
    """测试 4: 特殊字符搜索"""
    print("\n" + "=" * 60)
    print("测试 4: 特殊字符搜索")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入以下特殊字符组合:")
    print("   - '@#$%'")
    print("   - 'test@example.com'")
    print("   - 'C:\\\\Users\\\\test'")
    print("   - '(parenthesis)'")
    print("   - '[brackets]'")
    print("   - '{braces}'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 测试 1: 特殊符号
        print("\n--- 测试 4.1: 特殊符号 ---")
        success1, _ = await select_text(
            workspace,
            document_uri,
            search_text="@#$%",
            wait_seconds=2,
        )

        # 测试 2: 邮箱格式
        print("\n--- 测试 4.2: 邮箱格式 ---")
        success2, _ = await select_text(
            workspace,
            document_uri,
            search_text="test@example.com",
            wait_seconds=2,
        )

        # 测试 3: 括号
        print("\n--- 测试 4.3: 括号 ---")
        success3, _ = await select_text(
            workspace,
            document_uri,
            search_text="(parenthesis)",
            wait_seconds=2,
        )

        # 验证结果
        print("\n🔍 验证测试结果:")
        all_passed = True

        if success1:
            print("   ✓ 特殊符号 '@#$%' 搜索成功")
        else:
            print("   ❌ 特殊符号 '@#$%' 搜索失败")
            all_passed = False

        if success2:
            print("   ✓ 邮箱格式 'test@example.com' 搜索成功")
        else:
            print("   ❌ 邮箱格式 'test@example.com' 搜索失败")
            all_passed = False

        if success3:
            print("   ✓ 括号 '(parenthesis)' 搜索成功")
        else:
            print("   ❌ 括号 '(parenthesis)' 搜索失败")
            all_passed = False

        if all_passed:
            print("\n✅ 测试通过")
        else:
            print("\n⚠️  部分测试失败（请确认文档中包含相应文本）")


async def test_5_very_long_text() -> None:
    """测试 5: 长文本搜索"""
    print("\n" + "=" * 60)
    print("测试 5: 长文本搜索")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入一段很长的文本")

    long_text = "This is a very long text that " * 10  # 重复 10 次

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        success, data, error = await select_text(
            workspace,
            document_uri,
            search_text=long_text,
        )

        # 验证结果
        print("\n🔍 验证测试结果:")
        if success:
            print("   ✓ 长文本搜索成功")
            if data:
                print(f"   ✓ 找到 {data.get('matchCount', 0)} 个匹配")
            print("\n✅ 测试通过")
        else:
            print("   ❌ 长文本搜索失败")
            print("   ⚠️  请确认文档中包含该长文本")
            print("\n❌ 测试失败")


async def main() -> None:
    if len(sys.argv) < 2 or sys.argv[1] != "--test":
        print("Usage: python test_edge_cases.py --test <1-5|all>")
        return

    test_arg = sys.argv[2] if len(sys.argv) > 2 else "1"

    tests = {
        "1": test_1_no_matches,
        "2": test_2_empty_search_text,
        "3": test_3_out_of_bounds_index,
        "4": test_4_special_characters,
        "5": test_5_very_long_text,
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
