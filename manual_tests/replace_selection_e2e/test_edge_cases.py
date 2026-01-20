"""
Edge Cases Test

测试 word:replace:selection 的边界情况和错误处理。

测试场景:
1. 空选择替换（预期失败）
2. 替换为空字符串（删除选中内容）
3. 替换为包含图片的内容

Usage:
    # 运行单个测试
    uv run python manual_tests/replace_selection_e2e/test_edge_cases.py --test 1

    # 运行全部测试
    uv run python manual_tests/replace_selection_e2e/test_edge_cases.py --test all
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


async def test_1_replace_with_empty_selection() -> None:
    """
    测试场景 1: 空选择替换（预期返回错误）

    步骤:
        1. 启动 Workspace Socket.IO 服务器
        2. 等待 Word Add-In 连接
        3. 确保 Word 中没有选中任何文本（或选择已取消）
        4. 发送 word:replace:selection 请求
        5. 验证返回错误码 3002 (SELECTION_EMPTY)

    预期结果:
        - 返回 success=False
        - 错误码: 3002
        - 错误信息: "Current selection is empty"
    """
    print("\n" + "=" * 60)
    print("测试 1: 空选择替换（预期失败）")
    print("=" * 60)

    async with workspace_context() as workspace:
        print("\n⏳ 等待 Word Add-In 连接...")
        print("   请在 Word 中打开测试文档并确保 Add-In 已连接")
        print("   ⚠️  重要: 请确保 Word 中没有选中任何文本！")

        # Wait longer for user to deselect text
        await asyncio.sleep(5)

        # Get document URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"✅ 已连接文档: {document_uri}")

        content = {"text": "This should fail"}
        success, error = await replace_selection(
            workspace, document_uri, content, wait_seconds=2, return_error=True
        )

        print("\n📥 响应:")
        print(f"   Success: {success}")
        if error:
            print(f"   Error: {error}")

        print("\n🔍 验证:")
        if not success:
            if error and "3002" in error:
                print("   ✅ 测试通过: 正确返回错误码 3002 (SELECTION_EMPTY)")
            else:
                print(f"   ⚠️  测试警告: 错误码不匹配，预期 3002")
                print(f"   实际错误: {error}")
        else:
            print("   ⚠️  测试警告: 预期返回错误，但请求成功")
            print("   可能原因: Word 中有选中的内容")


async def test_2_replace_with_empty_string() -> None:
    """
    测试场景 2: 替换为空字符串（删除选中内容）

    步骤:
        1. 在 Word 中选中一些文本
        2. 发送 word:replace:selection 请求，content.text 为空字符串
        3. 验证选中内容被删除

    预期结果:
        - 选中的内容被删除（替换为空）
        - 返回 replaced=True, characterCount=0
    """
    print("\n" + "=" * 60)
    print("测试 2: 替换为空字符串（删除选中内容）")
    print("=" * 60)

    async with workspace_context() as workspace:
        print("\n⏳ 等待 Word Add-In 连接...")
        print("   请在 Word 中选中一些文本，这些文本将被删除")

        # Wait for connection
        if not await wait_for_connection(workspace):
            return

        # Get document URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"✅ 已连接文档: {document_uri}")

        content = {"text": ""}
        success, error = await replace_selection(
            workspace, document_uri, content, return_error=True
        )

        print("\n📥 响应:")
        print(f"   Success: {success}")
        if error:
            print(f"   Error: {error}")

        print("\n🔍 验证:")
        if success:
            print("   ✅ 测试通过: 选中内容已删除")
            print("   👀 请检查 Word 中的文本是否已被删除")
        else:
            print("   ⚠️  测试信息: 空字符串可能被拒绝")
            print("   这是正常行为，某些实现可能不允许空内容")
            print(f"   错误: {error}")


async def test_3_replace_with_image() -> None:
    """
    测试场景 3: 替换为包含图片的内容

    步骤:
        1. 在 Word 中选中一些文本
        2. 发送 word:replace:selection 请求，content.images 包含 base64 编码的图片
        3. 验证选中内容被图片替换

    预期结果:
        - 选中的内容被图片替换
        - 返回 replaced=True
    """
    print("\n" + "=" * 60)
    print("测试 3: 替换为包含图片的内容")
    print("=" * 60)

    async with workspace_context() as workspace:
        print("\n⏳ 等待 Word Add-In 连接...")
        print("   请在 Word 中选中一些文本")

        # Wait for connection
        if not await wait_for_connection(workspace):
            return

        # Get document URI
        document_uri = get_document_uri(workspace)
        if not document_uri:
            return

        print(f"✅ 已连接文档: {document_uri}")

        # Create a simple 1x1 red pixel PNG image (base64 encoded)
        base64_image = (
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=="
        )

        content = {
            "images": [
                {
                    "base64": base64_image,
                    "width": 100,
                    "height": 100,
                    "altText": "Test Image",
                }
            ]
        }
        success, error = await replace_selection(
            workspace, document_uri, content, return_error=True
        )

        print("\n📥 响应:")
        print(f"   Success: {success}")
        if error:
            print(f"   Error: {error}")

        print("\n🔍 验证:")
        if success:
            print("   ✅ 测试通过: 图片替换成功")
            print("   👀 请检查 Word 中是否显示了 100x100 的图片")
        else:
            print("   ⚠️  测试信息: 图片替换可能尚未实现")
            if error and "1003" in error:
                print("   Error: NOT_IMPLEMENTED - 图片替换功能待实现")


# ==============================================================================
# 测试运行器
# ==============================================================================


async def main() -> None:
    """主函数：运行指定的测试"""
    import argparse

    parser = argparse.ArgumentParser(description="Edge Cases E2E Tests")
    parser.add_argument(
        "--test",
        type=str,
        required=True,
        help="Test number to run (1-3) or 'all'",
    )
    args = parser.parse_args()

    tests = {
        "1": ("空选择替换（预期失败）", test_1_replace_with_empty_selection),
        "2": ("替换为空字符串（删除选中内容）", test_2_replace_with_empty_string),
        "3": ("替换为包含图片的内容", test_3_replace_with_image),
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
