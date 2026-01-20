"""
Basic Get Document Stats Test

测试文档统计信息获取功能。

测试场景:
1. 空白文档统计 - 获取空白文档的统计信息
2. 简单文档统计 - 获取包含简单文本的文档统计
3. 复杂文档统计 - 获取包含多种元素的文档统计
4. 大文档统计 - 获取包含大量内容的文档统计（性能测试）
"""

import asyncio
import sys

from manual_tests.test_helpers import (
    get_document_uri,
    wait_for_connection,
    workspace_context,
)
from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

# ==============================================================================
# 辅助函数和上下文管理器
# ==============================================================================


async def get_document_stats(
    workspace: OfficeWorkspace,
    document_uri: str,
    test_name: str = "",
) -> dict | None:
    """
    执行获取文档统计动作

    Args:
        workspace: Workspace 实例
        document_uri: 目标文档 URI
        test_name: 测试名称（用于提示）

    Returns:
        Optional[dict]: 文档统计数据，如果失败则返回 None
    """
    if test_name:
        print(f"\n📝 {test_name}")

    print("   提示: 请确保 Word 文档已打开并连接")

    action = OfficeAction(
        category="word",
        action_name="get:documentStats",
        params={
            "document_uri": document_uri,
        },
    )

    result = await workspace.execute(action)

    # 验证结果
    print("\n📊 验证结果:")
    if result.success:
        print("✅ 获取成功")
        if result.data:
            print(f"   字数: {result.data.get('wordCount', 'N/A')}")
            print(f"   字符数: {result.data.get('characterCount', 'N/A')}")
            print(f"   段落数: {result.data.get('paragraphCount', 'N/A')}")
        return result.data
    else:
        print(f"❌ 获取失败: {result.error}")
        return None


async def run_test_template(
    test_name: str,
    test_number: int,
    expected_description: str,
) -> bool:
    """
    测试执行模板：封装通用的测试流程

    Args:
        test_name: 测试名称
        test_number: 测试编号
        expected_description: 预期结果的描述

    Returns:
        bool: 测试是否成功
    """
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_name}")
    print("=" * 70)
    print(f"📋 预期结果: {expected_description}")

    async with workspace_context() as workspace:
        try:
            print("✅ Workspace 启动成功")

            # 等待连接
            if not await wait_for_connection(workspace):
                return False

            # 获取文档
            document_uri = get_document_uri(workspace)
            if not document_uri:
                return False

            print(f"✅ 使用文档: {document_uri}")

            # 执行获取文档统计
            stats = await get_document_stats(
                workspace,
                document_uri,
                "获取文档统计...",
            )

            if stats is None:
                return False

            print("\n" + "=" * 70)
            print(f"✅ 测试 {test_number} 完成")
            print("=" * 70)
            return True

        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback

            traceback.print_exc()
            return False


# ==============================================================================
# 测试函数（使用模板简化）
# ==============================================================================


async def test_empty_document_stats():
    """测试 1: 空白文档统计"""
    return await run_test_template(
        test_name="空白文档统计",
        test_number=1,
        expected_description="空白文档应该有 0 字、0 字符、0 段落（或 1 个空段落）",
    )


async def test_simple_document_stats():
    """测试 2: 简单文档统计"""
    return await run_test_template(
        test_name="简单文档统计",
        test_number=2,
        expected_description="简单文本文档应该正确统计字数、字符数和段落数",
    )


async def test_complex_document_stats():
    """测试 3: 复杂文档统计"""
    return await run_test_template(
        test_name="复杂文档统计",
        test_number=3,
        expected_description="包含多种元素的文档应该正确统计字数、字符数和段落数",
    )


async def test_large_document_stats():
    """测试 4: 大文档统计"""
    return await run_test_template(
        test_name="大文档统计",
        test_number=4,
        expected_description="大文档应该正确统计大量字数（性能测试）",
    )


async def run_all_tests():
    """运行所有基本统计获取测试"""
    print("\n🚀 运行所有文档统计获取测试...\n")
    results = []

    # 测试 1: 需要准备空白文档
    print("\n⚠️  请准备: 空白 Word 文档")
    input("按回车继续...")
    results.append(await test_empty_document_stats())
    await asyncio.sleep(2)

    # 测试 2: 需要准备简单文档
    print("\n⚠️  请准备: 包含几段简单文本的 Word 文档")
    input("按回车继续...")
    results.append(await test_simple_document_stats())
    await asyncio.sleep(2)

    # 测试 3: 需要准备复杂文档
    print("\n⚠️  请准备: 包含表格和图片的 Word 文档")
    input("按回车继续...")
    results.append(await test_complex_document_stats())
    await asyncio.sleep(2)

    # 测试 4: 需要准备长文档
    print("\n⚠️  请准备: 包含大量内容的 Word 文档（50+ 页）")
    input("按回车继续...")
    results.append(await test_large_document_stats())

    success = all(results)

    print("\n" + "=" * 70)
    print(f"📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    print("=" * 70)
    return success


# ==============================================================================
# 主程序入口
# ==============================================================================

# 测试映射表：用于命令行参数路由
TEST_MAPPING = {
    "1": test_empty_document_stats,
    "2": test_simple_document_stats,
    "3": test_complex_document_stats,
    "4": test_large_document_stats,
    "all": run_all_tests,
}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Basic Get Document Stats E2E Tests")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "all"],
        default="1",
        help="Test to run: 1=empty, 2=simple, 3=complex, 4=large, all=all tests",
    )

    args = parser.parse_args()

    try:
        test_func = TEST_MAPPING[args.test]
        success = asyncio.run(test_func())
        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
