"""
Selection Mode Tests

测试 word:select:text 的选择模式功能。

Usage:
    uv run python manual_tests/select_text_e2e/test_selection_modes.py --test 1
    uv run python manual_tests/select_text_e2e/test_selection_modes.py --test all
"""

import asyncio
import sys

from manual_tests.test_helpers import (
    get_document_uri,
    select_text,
    wait_for_connection,
    workspace_context,
)


async def test_1_select_mode() -> None:
    """测试 1: select 模式（高亮选区）"""
    print("\n" + "=" * 60)
    print("测试 1: select 模式（高亮选区）")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入多次 'Selection Test'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        success = await select_text(
            workspace,
            document_uri,
            search_text="Selection Test",
            selection_mode="select",
        )

        if success:
            print("\n✅ 测试完成: 检查 Word")
            print("   ✓ 'Selection Test' 应该被高亮选中")
            print("   ✓ 光标应该在选中文本上")
        else:
            print("\n❌ 测试失败: 选择操作未成功执行")


async def test_2_start_mode() -> None:
    """测试 2: start 模式（光标定位到开头）"""
    print("\n" + "=" * 60)
    print("测试 2: start 模式（光标定位到开头）")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入多次 'CursorPosition'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        success = await select_text(
            workspace,
            document_uri,
            search_text="CursorPosition",
            selection_mode="start",
        )

        if success:
            print("\n✅ 测试完成: 检查 Word")
            print("   ✓ 文本不应该被选中")
            print("   ✓ 光标应该在 'CursorPosition' 的开头")
            print("   ✓ 尝试输入文字，应该在开头插入")
        else:
            print("\n❌ 测试失败: 选择操作未成功执行")


async def test_3_end_mode() -> None:
    """测试 3: end 模式（光标定位到结尾）"""
    print("\n" + "=" * 60)
    print("测试 3: end 模式（光标定位到结尾）")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入多次 'EndPosition'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        success = await select_text(
            workspace,
            document_uri,
            search_text="EndPosition",
            selection_mode="end",
        )

        if success:
            print("\n✅ 测试完成: 检查 Word")
            print("   ✓ 文本不应该被选中")
            print("   ✓ 光标应该在 'EndPosition' 的结尾")
            print("   ✓ 尝试输入文字，应该在结尾追加")
        else:
            print("\n❌ 测试失败: 选择操作未成功执行")


async def test_4_mode_switching() -> None:
    """测试 4: 模式切换验证"""
    print("\n" + "=" * 60)
    print("测试 4: 模式切换验证")
    print("=" * 60)
    print("\n📋 准备: 在 Word 中输入多次 'ModeSwitch'")

    async with workspace_context() as workspace:
        if not await wait_for_connection(workspace):
            return

        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"\n✅ 使用文档: {document_uri}")

        # 第一次: select 模式
        print("\n--- 第 1 次选择: select 模式 ---")
        success1 = await select_text(
            workspace,
            document_uri,
            search_text="ModeSwitch",
            selection_mode="select",
            wait_seconds=2,
        )
        print("   检查: 文本应该被选中")

        # 第二次: start 模式
        print("\n--- 第 2 次选择: start 模式 ---")
        success2 = await select_text(
            workspace,
            document_uri,
            search_text="ModeSwitch",
            selection_mode="start",
            wait_seconds=2,
        )
        print("   检查: 光标应该在开头，文本不被选中")

        # 第三次: end 模式
        print("\n--- 第 3 次选择: end 模式 ---")
        success3 = await select_text(
            workspace,
            document_uri,
            search_text="ModeSwitch",
            selection_mode="end",
            wait_seconds=2,
        )
        print("   检查: 光标应该在结尾，文本不被选中")

        # 检查所有操作是否成功
        if success1 and success2 and success3:
            print("\n✅ 测试完成: 验证模式可以正确切换")
        else:
            print("\n❌ 测试失败: 部分选择操作未成功执行")
            if not success1:
                print("   - 第 1 次 select 模式失败")
            if not success2:
                print("   - 第 2 次 start 模式失败")
            if not success3:
                print("   - 第 3 次 end 模式失败")


async def main() -> None:
    if len(sys.argv) < 2 or sys.argv[1] != "--test":
        print("Usage: python test_selection_modes.py --test <1-4|all>")
        return

    test_arg = sys.argv[2] if len(sys.argv) > 2 else "1"

    tests = {
        "1": test_1_select_mode,
        "2": test_2_start_mode,
        "3": test_3_end_mode,
        "4": test_4_mode_switching,
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
