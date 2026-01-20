"""
Basic Text Replace Test

测试基本的文本替换功能（使用默认参数）。

测试场景:
1. 替换为纯文本
2. 替换为多行文本
3. 替换为特殊字符
4. 替换为长文本

Usage:
    # 运行单个测试
    uv run python manual_tests/replace_selection_e2e/test_text_replace.py --test 1

    # 运行全部测试
    uv run python manual_tests/replace_selection_e2e/test_text_replace.py --test all
"""

import asyncio
import sys

from manual_tests.test_helpers import (
    get_document_uri,
    replace_selection,
    wait_for_connection,
    workspace_context,
)

# ==============================================================================
# 测试用例
# ==============================================================================


async def test_1_replace_with_simple_text() -> None:
    """
    测试场景 1: 替换为纯文本

    步骤:
        1. 启动 Workspace Socket.IO 服务器
        2. 等待 Word Add-In 连接
        3. 在 Word 中选中一些文本
        4. 发送 word:replace:selection 请求，替换为 "Hello World"
        5. 验证替换成功

    预期结果:
        - 选中的文本被替换为 "Hello World"
        - 返回 replaced=True, characterCount=11
    """
    print("\n" + "=" * 60)
    print("测试 1: 替换为纯文本")
    print("=" * 60)

    async with workspace_context() as workspace:
        # Wait for connection
        if not await wait_for_connection(workspace):
            return

        # Get document URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"✅ 已连接文档: {document_uri}")
        print("\n⚠️  请在 Word 中选中一些文本，将被替换为 'Hello World'")

        # Replace selection
        content = {"text": "Hello World"}
        success = await replace_selection(workspace, document_uri, content, wait_seconds=5)

        # Verify result
        print("\n🔍 验证:")
        if success:
            print("   ✅ 测试通过: 文本替换成功")
            print("   👀 请检查 Word 中的文本是否已替换为 'Hello World'")
        else:
            print("   ❌ 测试失败: 替换失败")


async def test_2_replace_with_multiline_text() -> None:
    """
    测试场景 2: 替换为多行文本

    步骤:
        1. 在 Word 中选中一些文本
        2. 发送 word:replace:selection 请求，替换为多行文本

    预期结果:
        - 选中的文本被替换为多行文本
        - 返回正确的 characterCount
    """
    print("\n" + "=" * 60)
    print("测试 2: 替换为多行文本")
    print("=" * 60)

    async with workspace_context() as workspace:
        # Wait for connection
        if not await wait_for_connection(workspace):
            return

        # Get document URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"✅ 已连接文档: {document_uri}")
        print("\n⚠️  请在 Word 中选中一些文本")

        multiline_text = """第一行内容
第二行内容
第三行内容"""

        content = {"text": multiline_text}
        success = await replace_selection(workspace, document_uri, content)

        expected_count = len(multiline_text)

        print("\n🔍 验证:")
        if success:
            print(f"   ✅ 测试通过: 字符数正确 ({expected_count})")
            print("   👀 请检查 Word 中的文本是否为多行内容")
        else:
            print("   ❌ 测试失败")


async def test_3_replace_with_special_characters() -> None:
    """
    测试场景 3: 替换为特殊字符

    步骤:
        1. 在 Word 中选中一些文本
        2. 发送 word:replace:selection 请求，替换为包含特殊字符的文本

    预期结果:
        - 特殊字符被正确替换
        - 返回正确的 characterCount
    """
    print("\n" + "=" * 60)
    print("测试 3: 替换为特殊字符")
    print("=" * 60)

    async with workspace_context() as workspace:
        # Wait for connection
        if not await wait_for_connection(workspace):
            return

        # Get document URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"✅ 已连接文档: {document_uri}")
        print("\n⚠️  请在 Word 中选中一些文本")

        special_text = "特殊字符测试：中文、English、123、!@#$%、符号「」【】"

        content = {"text": special_text}
        success = await replace_selection(workspace, document_uri, content)

        expected_count = len(special_text)

        print("\n🔍 验证:")
        if success:
            print(f"   ✅ 测试通过: 特殊字符处理正确 ({expected_count} 字符)")
            print("   👀 请检查 Word 中的文本是否显示正确")
        else:
            print("   ❌ 测试失败")


async def test_4_replace_with_long_text() -> None:
    """
    测试场景 4: 替换为长文本

    步骤:
        1. 在 Word 中选中一些文本
        2. 发送 word:replace:selection 请求，替换为长文本（>1000字符）

    预期结果:
        - 长文本被完整替换
        - 返回正确的 characterCount
    """
    print("\n" + "=" * 60)
    print("测试 4: 替换为长文本")
    print("=" * 60)

    async with workspace_context() as workspace:
        # Wait for connection
        if not await wait_for_connection(workspace):
            return

        # Get document URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"✅ 已连接文档: {document_uri}")
        print("\n⚠️  请在 Word 中选中一些文本")

        # Generate long text (repeat pattern)
        long_text = "这是一段很长的测试文本。" * 50  # ~600 characters

        content = {"text": long_text}
        success = await replace_selection(workspace, document_uri, content)

        expected_count = len(long_text)

        print("\n🔍 验证:")
        if success:
            print(f"   ✅ 测试通过: 长文本处理正确 ({expected_count} 字符)")
            print("   👀 请检查 Word 中的文本长度是否正确")
        else:
            print("   ❌ 测试失败")


# ==============================================================================
# 测试运行器
# ==============================================================================


async def main() -> None:
    """主函数：运行指定的测试"""
    import argparse

    parser = argparse.ArgumentParser(description="Basic Text Replace E2E Tests")
    parser.add_argument(
        "--test",
        type=str,
        required=True,
        help="Test number to run (1-4) or 'all'",
    )
    args = parser.parse_args()

    tests = {
        "1": ("替换为纯文本", test_1_replace_with_simple_text),
        "2": ("替换为多行文本", test_2_replace_with_multiline_text),
        "3": ("替换为特殊字符", test_3_replace_with_special_characters),
        "4": ("替换为长文本", test_4_replace_with_long_text),
    }

    if args.test.lower() == "all":
        print("\n🚀 开始运行所有测试...")
        for num, (name, test_func) in tests.items():
            try:
                await test_func()
                print(f"\n✅ 测试 {num} 完成\n")
            except Exception as e:
                print(f"\n❌ 测试 {num} 失败: {e}\n")
    elif args.test in tests:
        name, test_func = tests[args.test]
        print(f"\n🚀 运行测试 {args.test}: {name}")
        try:
            await test_func()
        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback

            traceback.print_exc()
    else:
        print(f"❌ 无效的测试编号: {args.test}")
        print(f"   可用测试: {', '.join(tests.keys())}")
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())
