"""
Get Styles E2E Tests

测试 word:get:styles 功能的各种参数组合。

测试场景:
1. 获取所有正在使用的样式（默认参数）
2. 仅获取内置样式
3. 仅获取自定义样式
4. 获取包含详细信息的样式
5. 获取所有样式（包括未使用的）
"""

import asyncio
import sys
from collections.abc import AsyncGenerator
from contextlib import asynccontextmanager
from typing import Any

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

# ==============================================================================
# 辅助函数和上下文管理器
# ==============================================================================


@asynccontextmanager
async def workspace_context(host: str = "127.0.0.1", port: int = 3000) -> AsyncGenerator[OfficeWorkspace, None]:
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
    print("✅ Add-In 已连接")
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
    print(f"✅ 使用文档: {documents[0]}")
    return documents[0]


async def get_styles(
    workspace: OfficeWorkspace,
    document_uri: str,
    options: dict[str, Any] | None = None,
    wait_seconds: int = 2,
) -> dict[str, Any] | None:
    """
    执行获取样式动作

    Args:
        workspace: Workspace 实例
        document_uri: 目标文档 URI
        options: 样式获取选项
        wait_seconds: 执行前等待秒数

    Returns:
        Optional[dict]: 返回的样式数据，失败返回 None
    """
    print("\n📋 获取样式...")
    if options:
        print(f"   选项: {options}")

    await asyncio.sleep(wait_seconds)

    action = OfficeAction(
        category="word",
        action_name="get:styles",
        params={
            "document_uri": document_uri,
            **({"options": options} if options else {}),
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


def display_styles(styles: list[dict[str, Any]], detailed: bool = False) -> None:
    """
    显示样式列表

    Args:
        styles: 样式列表
        detailed: 是否显示详细信息
    """
    if not styles:
        print("   ⚠️  未找到样式")
        return

    print(f"\n   📚 样式列表 (共 {len(styles)} 个):")

    # 按类型分组
    by_type: dict[str, list[dict[str, Any]]] = {
        "Paragraph": [],
        "Character": [],
        "Table": [],
        "List": [],
    }

    for style in styles:
        style_type = style.get("type", "Unknown")
        if style_type in by_type:
            by_type[style_type].append(style)

    # 显示各类型样式
    for style_type, type_styles in by_type.items():
        if not type_styles:
            continue

        print(f"\n   {style_type} 样式 ({len(type_styles)} 个):")
        for style in type_styles[:10]:  # 最多显示 10 个
            name = style.get("name", "Unknown")
            built_in = style.get("builtIn", False)
            in_use = style.get("inUse", False)
            description = style.get("description")

            status = []
            if built_in:
                status.append("内置")
            else:
                status.append("自定义")
            if in_use:
                status.append("使用中")

            print(f"      • {name} [{', '.join(status)}]")

            if detailed and description:
                print(f"        描述: {description}")

        if len(type_styles) > 10:
            print(f"      ... 还有 {len(type_styles) - 10} 个")


async def run_test_template(
    test_name: str,
    test_number: int,
    options: dict[str, Any] | None = None,
    detailed_display: bool = False,
) -> bool:
    """
    测试执行模板：封装通用的测试流程

    Args:
        test_name: 测试名称
        test_number: 测试编号
        options: 样式获取选项
        detailed_display: 是否详细显示

    Returns:
        bool: 测试是否成功
    """
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_name}")
    print("=" * 70)

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
            data = await get_styles(workspace, document_uri, options)
            if data is None:
                return False

            # 显示结果
            styles = data.get("styles", [])
            display_styles(styles, detailed=detailed_display)

            # 统计信息
            print("\n   📈 统计信息:")
            print(f"   - 样式总数: {len(styles)}")

            by_type: dict[str, int] = {"Paragraph": 0, "Character": 0, "Table": 0, "List": 0}
            for style in styles:
                style_type = style.get("type", "Unknown")
                if style_type in by_type:
                    by_type[style_type] += 1

            for style_type, count in by_type.items():
                if count > 0:
                    print(f"   - {style_type}: {count} 个")

            built_in_count = sum(1 for s in styles if s.get("builtIn", False))
            custom_count = len(styles) - built_in_count
            in_use_count = sum(1 for s in styles if s.get("inUse", False))

            print(f"   - 内置样式: {built_in_count} 个")
            print(f"   - 自定义样式: {custom_count} 个")
            print(f"   - 使用中: {in_use_count} 个")

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


async def test_get_all_styles_in_use() -> bool:
    """测试 1: 获取所有正在使用的样式（默认参数）"""
    return await run_test_template(
        test_name="获取所有正在使用的样式（默认参数）",
        test_number=1,
        options=None,
        detailed_display=False,
    )


async def test_get_built_in_styles_only() -> bool:
    """测试 2: 仅获取内置样式"""
    return await run_test_template(
        test_name="仅获取内置样式",
        test_number=2,
        options={
            "includeBuiltIn": True,
            "includeCustom": False,
            "includeUnused": False,
            "detailedInfo": False,
        },
        detailed_display=False,
    )


async def test_get_custom_styles_only() -> bool:
    """测试 3: 仅获取自定义样式"""
    return await run_test_template(
        test_name="仅获取自定义样式",
        test_number=3,
        options={
            "includeBuiltIn": False,
            "includeCustom": True,
            "includeUnused": True,  # 包括未使用的自定义样式
            "detailedInfo": False,
        },
        detailed_display=False,
    )


async def test_get_styles_with_detailed_info() -> bool:
    """测试 4: 获取包含详细信息的样式"""
    return await run_test_template(
        test_name="获取包含详细信息的样式",
        test_number=4,
        options={
            "includeBuiltIn": True,
            "includeCustom": True,
            "includeUnused": False,
            "detailedInfo": True,  # 包含描述信息
        },
        detailed_display=True,
    )


async def test_get_all_styles_including_unused() -> bool:
    """测试 5: 获取所有样式（包括未使用的）"""
    return await run_test_template(
        test_name="获取所有样式（包括未使用的）",
        test_number=5,
        options={
            "includeBuiltIn": True,
            "includeCustom": True,
            "includeUnused": True,  # 包括未使用的样式
            "detailedInfo": False,
        },
        detailed_display=False,
    )


async def run_all_tests() -> bool:
    """运行所有样式获取测试"""
    print("\n🚀 运行所有样式获取测试...\n")
    results = []
    results.append(await test_get_all_styles_in_use())
    await asyncio.sleep(2)
    results.append(await test_get_built_in_styles_only())
    await asyncio.sleep(2)
    results.append(await test_get_custom_styles_only())
    await asyncio.sleep(2)
    results.append(await test_get_styles_with_detailed_info())
    await asyncio.sleep(2)
    results.append(await test_get_all_styles_including_unused())
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
    "1": test_get_all_styles_in_use,
    "2": test_get_built_in_styles_only,
    "3": test_get_custom_styles_only,
    "4": test_get_styles_with_detailed_info,
    "5": test_get_all_styles_including_unused,
    "all": run_all_tests,
}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Get Styles E2E Tests")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "5", "all"],
        default="1",
        help="Test to run: 1=all in use, 2=built-in only, 3=custom only, 4=detailed, 5=including unused, all=all tests",
    )

    args = parser.parse_args()

    try:
        test_func = TEST_MAPPING[args.test]
        success = asyncio.run(test_func())
        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
