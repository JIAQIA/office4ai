"""
Edge Cases Get Visible Content E2E Tests

测试 word:get:visibleContent 的边界情况和异常场景。

测试场景:
1. 获取超长文档的可见内容
2. 获取包含特殊字符的内容
3. 获取包含嵌入对象的内容
4. 多次连续获取可见内容
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


async def get_visible_content(
    workspace: OfficeWorkspace,
    document_uri: str,
    options: dict[str, Any] | None = None,
    wait_seconds: int = 2,
) -> dict[str, Any] | None:
    """执行获取可见内容动作"""
    print("\n📋 获取可见内容...")
    if options:
        print(f"   选项: {options}")

    await asyncio.sleep(wait_seconds)

    action = OfficeAction(
        category="word",
        action_name="get:visibleContent",
        params={
            "document_uri": document_uri,
            **({"options": options} if options else {}),
        },
    )

    print(f"   发送动作: {action.category}:{action.action_name}")

    result = await workspace.execute(action)

    print("\n📊 验证结果:")
    if result.success:
        print("✅ 获取成功")
        return result.data
    else:
        print(f"❌ 获取失败: {result.error}")
        return None


def display_content_summary(data: dict[str, Any]) -> None:
    """显示内容摘要"""
    print("\n📄 可见内容摘要:")
    text = data.get("text", "")
    metadata = data.get("metadata", {})
    elements = data.get("elements", [])

    print(f"   文本长度: {len(text)} 字符")
    print(f"   字符数: {metadata.get('characterCount', 0)}")
    print(f"   元素数量: {len(elements)}")

    # 统计元素类型
    elem_types: dict[str, int] = {}
    for elem in elements:
        elem_type = elem.get("type", "unknown")
        elem_types[elem_type] = elem_types.get(elem_type, 0) + 1

    if elem_types:
        print("   元素类型分布:")
        for elem_type, count in elem_types.items():
            print(f"      - {elem_type}: {count}")


async def run_test_template(
    test_name: str,
    test_number: int,
    options: dict[str, Any] | None = None,
) -> bool:
    """测试执行模板"""
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_name}")
    print("=" * 70)

    async with workspace_context() as workspace:
        try:
            print("✅ Workspace 启动成功")

            if not await wait_for_connection(workspace):
                return False

            document_uri = get_document_uri(workspace)
            if not document_uri:
                return False

            data = await get_visible_content(workspace, document_uri, options)
            if data is None:
                return False

            display_content_summary(data)

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


async def test_very_long_document() -> bool:
    """测试 1: 获取超长文档的可见内容"""
    return await run_test_template(
        test_name="获取超长文档的可见内容",
        test_number=1,
        options={"maxTextLength": 10000},  # 设置较大的限制
    )


async def test_special_characters() -> bool:
    """测试 2: 获取包含特殊字符的内容"""
    return await run_test_template(
        test_name="获取包含特殊字符的内容",
        test_number=2,
        options={"includeText": True},
    )


async def test_embedded_objects() -> bool:
    """测试 3: 获取包含嵌入对象的内容"""
    return await run_test_template(
        test_name="获取包含嵌入对象的内容",
        test_number=3,
        options={"includeText": True, "includeImages": True, "includeTables": True},
    )


async def test_consecutive_requests() -> bool:
    """测试 4: 多次连续获取可见内容"""
    print("\n" + "=" * 70)
    print("🧪 测试 4: 多次连续获取可见内容")
    print("=" * 70)

    async with workspace_context() as workspace:
        try:
            print("✅ Workspace 启动成功")

            if not await wait_for_connection(workspace):
                return False

            document_uri = get_document_uri(workspace)
            if not document_uri:
                return False

            # 连续执行3次
            for i in range(1, 4):
                print(f"\n--- 第 {i} 次获取 ---")
                data = await get_visible_content(workspace, document_uri, None, wait_seconds=1)
                if data is None:
                    return False
                display_content_summary(data)
                await asyncio.sleep(1)

            print("\n" + "=" * 70)
            print("✅ 测试 4 完成 (连续3次)")
            print("=" * 70)
            return True

        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback

            traceback.print_exc()
            return False


async def run_all_tests() -> bool:
    """运行所有边界情况测试"""
    print("\n🚀 运行所有边界情况获取可见内容测试...\n")
    results = []
    results.append(await test_very_long_document())
    await asyncio.sleep(2)
    results.append(await test_special_characters())
    await asyncio.sleep(2)
    results.append(await test_embedded_objects())
    await asyncio.sleep(2)
    results.append(await test_consecutive_requests())
    success = all(results)

    print("\n" + "=" * 70)
    print(f"📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    print("=" * 70)
    return success


# ==============================================================================
# 主程序入口
# ==============================================================================

TEST_MAPPING = {
    "1": test_very_long_document,
    "2": test_special_characters,
    "3": test_embedded_objects,
    "4": test_consecutive_requests,
    "all": run_all_tests,
}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Get Visible Content E2E Tests - Edge Cases")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "all"],
        default="1",
        help="Test to run",
    )

    args = parser.parse_args()

    try:
        test_func = TEST_MAPPING[args.test]
        success = asyncio.run(test_func())
        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
