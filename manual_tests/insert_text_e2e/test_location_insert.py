"""
Location-based Insert Text Tests

测试不同插入位置（location 参数）的文本插入功能。

测试场景:
1. location="Cursor" - 在光标位置插入（默认）
2. location="Start" - 在文档开头插入
3. location="End" - 在文档末尾插入
4. 连续多次插入测试位置累积效果
"""

import asyncio
import sys
from dataclasses import dataclass

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from manual_tests.test_helpers import ready_workspace


@dataclass
class InsertTestConfig:
    """插入测试配置"""

    test_name: str
    text: str
    location: str
    prep_time: int = 3
    verify_message: str = ""
    description: str = ""


def print_test_header(test_number: int, test_name: str) -> None:
    """打印测试标题"""
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_name}")
    print("=" * 70)


def print_test_footer(test_number: int, success: bool) -> None:
    """打印测试完成信息"""
    status = "✅" if success else "❌"
    print(f"\n{status} 测试 {test_number} {'通过' if success else '失败'}")
    print("=" * 70)


async def execute_single_insert_test(
    workspace: OfficeWorkspace,
    document_uri: str,
    config: InsertTestConfig,
) -> bool:
    """
    执行单个插入测试

    Args:
        workspace: OfficeWorkspace 实例
        document_uri: 目标文档 URI
        config: 测试配置

    Returns:
        bool: 测试是否成功
    """
    print(f"\n📝 {config.description}")
    if config.verify_message:
        print(f"   {config.verify_message}")

    await asyncio.sleep(config.prep_time)

    action = OfficeAction(
        category="word",
        action_name="insert:text",
        params={
            "document_uri": document_uri,
            "text": config.text,
            "location": config.location,
        },
    )

    result = await workspace.execute(action)

    print("\n📊 验证结果:")
    if result.success:
        print("✅ 插入成功")
        print(f"   返回数据: {result.data}")
        return True
    else:
        print(f"❌ 插入失败: {result.error}")
        return False


async def run_single_test(test_number: int, config: InsertTestConfig) -> bool:
    """
    运行单个完整测试（包含 setup 和 teardown）

    Args:
        test_number: 测试编号
        config: 测试配置

    Returns:
        bool: 测试是否成功
    """
    print_test_header(test_number, config.test_name)

    try:
        async with ready_workspace() as (workspace, document_uri):
            success = await execute_single_insert_test(workspace, document_uri, config)
            print_test_footer(test_number, success)
            return success
    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback

        traceback.print_exc()
        print_test_footer(test_number, False)
        return False


# ============================================================================
# 测试场景实现
# ============================================================================


async def test_insert_at_cursor() -> bool:
    """
    测试 1: 在光标位置插入文本
    location="Cursor"（默认值）
    """
    config = InsertTestConfig(
        test_name="在光标位置插入文本",
        text="[光标位置文本]",
        location="Cursor",
        prep_time=5,
        description="在光标位置插入: '[光标位置文本]'",
        verify_message="⚠️  重要: 请先将光标移动到文档中的任意位置",
    )
    return await run_single_test(1, config)


async def test_insert_at_start() -> bool:
    """
    测试 2: 在文档开头插入文本
    location="Start"
    """
    config = InsertTestConfig(
        test_name="在文档开头插入文本",
        text="[文档开头标题]\n",
        location="Start",
        prep_time=3,
        description="在文档开头插入: '[文档开头标题]\\n'",
        verify_message="提示: 文本将插入到文档的最开始位置",
    )
    return await run_single_test(2, config)


async def test_insert_at_end() -> bool:
    """
    测试 3: 在文档末尾插入文本
    location="End"
    """
    config = InsertTestConfig(
        test_name="在文档末尾插入文本",
        text="\n[文档末尾追加内容]",
        location="End",
        prep_time=3,
        description="在文档末尾插入: '\\n[文档末尾追加内容]'",
        verify_message="提示: 文本将插入到文档的最后位置",
    )
    return await run_single_test(3, config)


async def test_multiple_insertions() -> bool:
    """
    测试 4: 连续多次插入测试
    验证不同位置的连续插入效果
    """
    print_test_header(4, "连续多次插入测试")

    async with workspace_test_context() as (workspace, document_uri):
        if not workspace or not document_uri:
            print_test_footer(4, False)
            return False

        # 连续插入三次，每次使用不同的位置
        inserts = [
            ("Start", "=== 第一次插入（开头） ===\n"),
            ("End", "=== 第二次插入（末尾） ===\n"),
            ("Cursor", "=== 第三次插入（光标） ==="),
        ]

        print("\n📝 将执行 3 次连续插入:")
        for i, (location, text) in enumerate(inserts, 1):
            print(f"   {i}. location={location}: {text.strip()}")

        print("\n   提示: 请将光标放在文档中间位置")
        await asyncio.sleep(5)

        results = []
        for i, (location, text) in enumerate(inserts, 1):
            print(f"\n--- 执行第 {i} 次插入 ---")

            action = OfficeAction(
                category="word",
                action_name="insert:text",
                params={
                    "document_uri": document_uri,
                    "text": text,
                    "location": location,
                },
            )

            result = await workspace.execute(action)
            results.append(result.success)

            if result.success:
                print(f"✅ 第 {i} 次插入成功")
            else:
                print(f"❌ 第 {i} 次插入失败: {result.error}")

            if i < len(inserts):
                await asyncio.sleep(1)

        print("\n📊 验证结果:")
        success_count = sum(results)
        print(f"   成功: {success_count}/{len(results)}")

        if all(results):
            print("\n   ✅ 所有插入都成功！")
            print("   请检查 Word 文档，确认三次插入的位置正确")
        else:
            print("\n   ⚠️  部分插入失败")

        success = all(results)
        print_test_footer(4, success)
        return success


async def run_all_tests() -> bool:
    """运行所有位置插入测试"""
    print("\n🚀 运行所有位置插入测试...\n")
    results = []
    results.append(await test_insert_at_cursor())
    await asyncio.sleep(2)
    results.append(await test_insert_at_start())
    await asyncio.sleep(2)
    results.append(await test_insert_at_end())
    await asyncio.sleep(2)
    results.append(await test_multiple_insertions())
    success = all(results)

    print("\n" + "=" * 70)
    print(f"📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    print("=" * 70)
    return success


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Location-based Insert Text E2E Tests")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "all"],
        default="1",
        help="Test to run: 1=cursor, 2=start, 3=end, 4=multiple, all=all tests",
    )

    args = parser.parse_args()

    try:
        if args.test == "1":
            success = asyncio.run(test_insert_at_cursor())
        elif args.test == "2":
            success = asyncio.run(test_insert_at_start())
        elif args.test == "3":
            success = asyncio.run(test_insert_at_end())
        elif args.test == "4":
            success = asyncio.run(test_multiple_insertions())
        else:  # all
            success = asyncio.run(run_all_tests())

        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
