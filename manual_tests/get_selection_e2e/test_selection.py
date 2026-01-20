"""
Get Selection E2E Tests

测试 word:get:selection 功能的各种选区状态。
"""

import asyncio
import sys
from typing import Any

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


async def get_selection(
    workspace: OfficeWorkspace,
    document_uri: str,
    wait_seconds: int = 2,
) -> dict[str, Any] | None:
    """
    执行获取选区动作

    Args:
        workspace: Workspace 实例
        document_uri: 目标文档 URI
        wait_seconds: 执行前等待秒数

    Returns:
        Optional[dict]: 返回的选区数据，失败返回 None
    """
    print("\n📍 获取选区信息...")
    await asyncio.sleep(wait_seconds)

    action = OfficeAction(
        category="word",
        action_name="get:selection",
        params={
            "document_uri": document_uri,
        },
    )

    print(f"   发送动作: {action.category}:{action.action_name}")

    result = await workspace.execute(action)

    # 验证结果
    print("\n📊 验证结果:")
    if result.success:
        print("✅ 获取成功")
        return result.data
    else:
        print(f"❌ 获取失败: {result.error}")
        return None


def display_selection(selection: dict[str, Any]) -> None:
    """
    显示选区信息

    Args:
        selection: 选区数据
    """
    is_empty = selection.get("isEmpty", False)
    selection_type = selection.get("type", "Unknown")
    start = selection.get("start")
    end = selection.get("end")
    text = selection.get("text")

    print("\n   📌 选区状态:")
    print(f"      是否为空: {is_empty}")
    print(f"      选区类型: {selection_type}")

    if start is not None:
        print(f"      起始位置: {start}")
    if end is not None:
        print(f"      结束位置: {end}")
    if text is not None:
        print(f"      选区文本: '{text}'")
        print(f"      文本长度: {len(text)} 字符")

    # 验证一致性
    if not is_empty and start is not None and end is not None and text is not None:
        expected_length = end - start
        actual_length = len(text)
        if expected_length == actual_length:
            print(f"      ✅ 长度一致: {actual_length} 字符")
        else:
            print(f"      ⚠️  长度不一致: 预期 {expected_length}, 实际 {actual_length}")


async def run_test_template(
    test_name: str,
    test_number: int,
    preparation: str,
) -> bool:
    """
    测试执行模板：封装通用的测试流程

    Args:
        test_name: 测试名称
        test_number: 测试编号
        preparation: 前置操作说明

    Returns:
        bool: 测试是否成功
    """
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_name}")
    print("=" * 70)

    print("\n📋 前置操作:")
    print(f"   {preparation}")

    input("\n⏸️  完成前置操作后按 Enter 继续...")

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

            # 执行获取
            data = await get_selection(workspace, document_uri)
            if data is None:
                return False

            # 显示结果
            display_selection(data)

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
# 测试函数
# ==============================================================================


async def test_get_normal_selection() -> bool:
    """测试 1: 获取正常选区（有高亮文本）"""
    return await run_test_template(
        test_name="获取正常选区（有高亮文本）",
        test_number=1,
        preparation="在 Word 文档中选中一段文本（创建高亮选区）",
    )


async def test_get_insertion_point() -> bool:
    """测试 2: 获取光标位置（无高亮文本）"""
    return await run_test_template(
        test_name="获取光标位置（无高亮文本）",
        test_number=2,
        preparation="在 Word 文档中点击放置光标，不要选中任何文本",
    )


async def test_get_no_selection() -> bool:
    """测试 3: 获取无选区状态"""
    return await run_test_template(
        test_name="获取无选区状态",
        test_number=3,
        preparation="确保文档没有任何选中内容（点击文档外部区域）",
    )


async def test_verify_lightweight() -> bool:
    """测试 4: 验证轻量级特性"""
    print("\n" + "=" * 70)
    print("🧪 测试 4: 验证轻量级特性")
    print("=" * 70)

    print("\n📋 此测试对比 word:get:selection 和 word:get:selectedContent 的性能")

    async with workspace_context() as workspace:
        try:
            print("✅ Workspace 启动成功")

            if not await wait_for_connection(workspace):
                return False

            document_uri = get_document_uri(workspace)
            if not document_uri:
                return False

            # 测试 word:get:selection
            print("\n⏱️  测试 word:get:selection 性能...")
            import time

            await asyncio.sleep(2)
            start_time = time.time()

            action1 = OfficeAction(
                category="word",
                action_name="get:selection",
                params={"document_uri": document_uri},
            )
            result1 = await workspace.execute(action1)

            selection_time = time.time() - start_time

            if result1.success:
                print(f"✅ word:get:selection 耗时: {selection_time:.3f} 秒")
            else:
                print(f"❌ word:get:selection 失败: {result1.error}")
                return False

            await asyncio.sleep(2)

            # 测试 word:get:selectedContent
            print("\n⏱️  测试 word:get:selectedContent 性能...")
            start_time = time.time()

            action2 = OfficeAction(
                category="word",
                action_name="get:selectedContent",
                params={"document_uri": document_uri},
            )
            result2 = await workspace.execute(action2)

            content_time = time.time() - start_time

            if result2.success:
                print(f"✅ word:get:selectedContent 耗时: {content_time:.3f} 秒")
            else:
                print(f"⚠️  word:get:selectedContent 失败: {result2.error}")
                print("   （继续测试）")

            # 对比结果
            print("\n📊 性能对比:")
            print(f"   word:get:selection:      {selection_time:.3f} 秒")
            print(f"   word:get:selectedContent: {content_time:.3f} 秒")

            if selection_time < content_time:
                speedup = content_time / selection_time
                print(f"   ✅ word:get:selection 快 {speedup:.1f}x")
            else:
                print("   ⚠️  word:get:selection 未能显示出性能优势")

            print("\n💡 结论: word:get:selection 是轻量级查询，适合快速获取位置信息")

            print("\n" + "=" * 70)
            print("✅ 测试 4 完成")
            print("=" * 70)
            return True

        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback

            traceback.print_exc()
            return False


async def run_all_tests() -> bool:
    """运行所有选区获取测试"""
    print("\n🚀 运行所有选区获取测试...\n")
    results = []
    results.append(await test_get_normal_selection())
    results.append(await test_get_insertion_point())
    results.append(await test_get_no_selection())
    results.append(await test_verify_lightweight())
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
    "1": test_get_normal_selection,
    "2": test_get_insertion_point,
    "3": test_get_no_selection,
    "4": test_verify_lightweight,
    "all": run_all_tests,
}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Get Selection E2E Tests")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "all"],
        default="1",
        help="Test to run: 1=normal selection, 2=insertion point, 3=no selection, 4=lightweight verify, all=all tests",
    )

    args = parser.parse_args()

    try:
        test_func = TEST_MAPPING[args.test]
        success = asyncio.run(test_func())
        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
