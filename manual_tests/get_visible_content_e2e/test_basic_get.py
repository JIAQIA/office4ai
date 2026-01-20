"""
Basic Get Visible Content E2E Tests

测试 word:get:visibleContent 的基础功能。

测试场景:
1. 获取当前可见区域的文本内容
2. 获取空文档的可见内容
3. 获取包含格式化文本的可见内容
4. 获取包含多种元素（文本、图片、表格）的可见内容
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
    """
    执行获取可见内容动作

    Args:
        workspace: Workspace 实例
        document_uri: 目标文档 URI
        options: 内容获取选项
        wait_seconds: 执行前等待秒数

    Returns:
        Optional[dict]: 返回的可见内容数据，失败返回 None
    """
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

    # 验证结果
    print("\n📊 验证结果:")
    if result.success:
        print("✅ 获取成功")
        return result.data
    else:
        print(f"❌ 获取失败: {result.error}")
        return None


def debug_text_chars(text: str, label: str = "文本字符详情") -> None:
    """
    调试文本字符，显示所有特殊字符

    Args:
        text: 要分析的文本
        label: 标签
    """
    print(f"\n   🔍 {label}:")
    print(f"      原始长度: {len(text)} 字符")
    print(f"      repr(): {repr(text)}")

    # 显示每个字符的详细信息
    if text:
        print("      字符详情:")
        for i, char in enumerate(text):
            code_point = ord(char)
            char_name = {
                10: "\\n (LF)",
                13: "\\r (CR)",
                9: "\\t (TAB)",
                32: "SPACE",
            }.get(code_point, "")

            # 控制字符范围
            if code_point < 32 or code_point == 127:
                print(f"         [{i}] = {code_point:3d} = {repr(char)} {char_name}")
            elif code_point in (8232, 8233):  # Unicode line separator
                print(f"         [{i}] = {code_point:3d} = U+{code_point:04X} (Unicode Line Separator)")
            else:
                # 可打印字符只显示前20个
                if i < 20:
                    print(f"         [{i}] = {code_point:3d} = {repr(char)}")


def display_visible_content(data: dict[str, Any], show_details: bool = False) -> None:
    """
    显示可见内容

    Args:
        data: 可见内容数据
        show_details: 是否显示详细信息（格式、元素内容等）
    """
    print("\n📄 可见内容:")
    metadata = data.get("metadata", {})
    text = data.get("text", "")

    # 显示文本预览（前50个字符）
    text_preview = text[:50] + "..." if len(text) > 50 else text
    print(f"   文本预览: {text_preview}")
    print(f"   字符数: {metadata.get('characterCount', 0)}")
    print(f"   是否为空: {metadata.get('isEmpty', False)}")

    # 调试：显示字符详情
    debug_text_chars(text, "完整文本字符分析")

    # 同时分析每个文本元素
    print("\n   📋 各元素文本分析:")
    elements = data.get("elements", [])
    for i, elem in enumerate(elements, 1):
        elem_type = elem.get("type", "unknown")
        if elem_type == "text":
            elem_content = elem.get("content", {})
            elem_text = elem_content.get("text", "")
            debug_text_chars(elem_text, f"元素 #{i} 文本")

    elements = data.get("elements", [])
    if elements:
        print(f"\n   元素列表 (共 {len(elements)} 个):")
        # 按从上到下的顺序遍历元素
        for i, elem in enumerate(elements, 1):
            elem_type = elem.get("type", "unknown")
            elem_content = elem.get("content", {})

            if elem_type == "text":
                # 打印文本内容
                text_content = elem_content.get("text", "")
                # 截断过长的文本
                display_text = text_content[:100] + "..." if len(text_content) > 100 else text_content
                print(f"      {i}. [Text] {display_text}")

            elif elem_type == "image":
                # 打印图片占位符
                width = elem_content.get("width", "?")
                height = elem_content.get("height", "?")
                alt_text = elem_content.get("altText", "")
                metadata_str = f"width={width}, height={height}"
                if alt_text:
                    metadata_str += f", alt={repr(alt_text[:30])}"
                print(f"      {i}. [Image]({metadata_str})")

            elif elem_type == "table":
                # 打印表格信息
                rows = elem_content.get("rows", "?")
                cols = elem_content.get("columns", "?")
                print(f"      {i}. [Table](rows={rows}, columns={cols})")

                # 如果需要详细信息，显示表格内容预览
                if show_details:
                    cells = elem_content.get("cells", [])
                    if cells:
                        print("         单元格预览 (前5个):")
                        for j, cell in enumerate(cells[:5], 1):
                            cell_text = cell.get("text", "")
                            print(f"            {j}. {repr(cell_text[:30])}")

            else:
                # 其他未知类型
                print(f"      {i}. [{elem_type}]")

        if len(elements) > 10:
            print(f"      ... 还有 {len(elements) - 10} 个")


async def run_test_template(
    test_name: str,
    test_number: int,
    options: dict[str, Any] | None = None,
) -> bool:
    """
    测试执行模板：封装通用的测试流程

    Args:
        test_name: 测试名称
        test_number: 测试编号
        options: 内容获取选项

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
            data = await get_visible_content(workspace, document_uri, options)
            if data is None:
                return False

            # 显示结果
            display_visible_content(data)

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


async def test_get_visible_text_only() -> bool:
    """测试 1: 获取当前可见区域的文本内容"""
    return await run_test_template(
        test_name="获取当前可见区域的文本内容",
        test_number=1,
        options=None,
    )


async def test_get_empty_document() -> bool:
    """测试 2: 获取空文档的可见内容"""
    print("\n" + "=" * 70)
    print("🧪 测试 2: 获取空文档的可见内容")
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

            # ⚠️ 前置检查：先获取当前内容，验证是否为空文档
            # 注意：Word 新建的空白文档通常包含段落标记等隐藏字符（约1-10字符）
            # 所以我们使用阈值判断：字符数 <= 10 视为"用户语义的空文档"
            print("\n⚠️  前置检查：验证文档是否为空...")
            initial_data = await get_visible_content(workspace, document_uri, None, wait_seconds=1)
            if initial_data is None:
                return False

            char_count = initial_data.get("metadata", {}).get("characterCount", 0)
            EMPTY_THRESHOLD = 10  # 空文档阈值：Word 新建文档通常有1-10个隐藏字符

            if char_count > EMPTY_THRESHOLD:
                print("\n" + "=" * 70)
                print("❌ 测试失败：文档不是空的！")
                print(f"   当前文档包含 {char_count} 个字符")
                print(f"   内容预览: {repr(initial_data.get('text', '')[:50])}")
                print("\n   ⚠️  请准备一个空文档（新建空白 Word 文档）后再运行此测试")
                print(f"   说明：字符数应 ≤ {EMPTY_THRESHOLD}（Word 新建文档通常有少量隐藏字符）")
                print("=" * 70)
                return False

            print(f"✅ 文档状态验证通过（字符数: {char_count} ≤ {EMPTY_THRESHOLD}，视为空文档）")

            # 执行测试
            data = await get_visible_content(workspace, document_uri, None)
            if data is None:
                return False

            # 显示结果
            display_visible_content(data)

            # 验证空文档的元数据（使用阈值判断）
            final_char_count = data.get("metadata", {}).get("characterCount", 0)
            assert final_char_count <= EMPTY_THRESHOLD, f"字符数应该 ≤ {EMPTY_THRESHOLD}，实际为 {final_char_count}"

            print("\n" + "=" * 70)
            print("✅ 测试 2 完成（空文档验证通过）")
            print("=" * 70)
            return True

        except AssertionError as e:
            print(f"\n❌ 断言失败: {e}")
            return False
        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback

            traceback.print_exc()
            return False


async def test_get_formatted_text() -> bool:
    """测试 3: 获取包含格式化文本的可见内容"""
    print("\n" + "=" * 70)
    print("🧪 测试 3: 获取包含格式化文本的可见内容")
    print("=" * 70)

    async with workspace_context() as workspace:
        try:
            print("✅ Workspace 启动成功")

            if not await wait_for_connection(workspace):
                return False

            document_uri = get_document_uri(workspace)
            if not document_uri:
                return False

            # 使用 detailedMetadata 选项获取详细格式信息
            print("\n💡 提示：请在 Word 中准备包含格式化的文本（粗体、斜体、不同字体等）")
            await asyncio.sleep(2)

            data = await get_visible_content(
                workspace,
                document_uri,
                {"detailedMetadata": True, "includeText": True},
            )
            if data is None:
                return False

            # 显示详细内容
            display_visible_content(data, show_details=True)

            print("\n📝 验证格式化信息:")
            elements = data.get("elements", [])

            # 协议规定的类型（Confluence: word:get:visibleContent）
            # AnyContentElement.type 应该是: 'text' | 'image' | 'table' | 'other'
            text_elements = [e for e in elements if e.get("type") == "text"]
            image_elements = [e for e in elements if e.get("type") == "image"]
            table_elements = [e for e in elements if e.get("type") == "table"]

            print("   元素类型统计:")
            print(f"   - 文本段落: {len(text_elements)} 个")
            print(f"   - 图片: {len(image_elements)} 个")
            print(f"   - 表格: {len(table_elements)} 个")

            if text_elements:
                print("\n   文本段落详情 (前3个):")
                for i, elem in enumerate(text_elements[:3], 1):
                    content = elem.get("content", {})
                    text = content.get("text", "")
                    # 尝试获取格式信息
                    format_info = content.get("format", {})
                    if format_info:
                        bold = "粗体" if format_info.get("bold") else ""
                        italic = "斜体" if format_info.get("italic") else ""
                        font_size = format_info.get("fontSize")
                        format_str = ", ".join(filter(None, [bold, italic, f"字号{font_size}" if font_size else None]))
                        print(f"      {i}. {repr(text[:50])}")
                        if format_str:
                            print(f"         格式: {format_str}")
                    else:
                        print(f"      {i}. {repr(text[:50])}")
            else:
                print("   ⚠️  未找到文本段落")

            print("\n" + "=" * 70)
            print("✅ 测试 3 完成")
            print("=" * 70)
            return True

        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback

            traceback.print_exc()
            return False


async def test_get_mixed_elements() -> bool:
    """测试 4: 获取包含多种元素的可见内容"""
    return await run_test_template(
        test_name="获取包含多种元素（文本、图片、表格）的可见内容",
        test_number=4,
        options={"includeText": True, "includeImages": True, "includeTables": True},
    )


async def run_all_tests() -> bool:
    """运行所有基础可见内容获取测试"""
    print("\n🚀 运行所有基础可见内容获取测试...\n")
    results = []
    results.append(await test_get_visible_text_only())
    await asyncio.sleep(2)
    results.append(await test_get_empty_document())
    await asyncio.sleep(2)
    results.append(await test_get_formatted_text())
    await asyncio.sleep(2)
    results.append(await test_get_mixed_elements())
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
    "1": test_get_visible_text_only,
    "2": test_get_empty_document,
    "3": test_get_formatted_text,
    "4": test_get_mixed_elements,
    "all": run_all_tests,
}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Get Visible Content E2E Tests - Basic")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "all"],
        default="1",
        help="Test to run: 1=text only, 2=empty document, 3=formatted text, 4=mixed elements, all=all tests",
    )

    args = parser.parse_args()

    try:
        test_func = TEST_MAPPING[args.test]
        success = asyncio.run(test_func())
        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
