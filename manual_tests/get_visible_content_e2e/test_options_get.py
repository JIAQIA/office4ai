"""
Options Get Visible Content E2E Tests

测试 word:get:visibleContent 的各种选项参数组合。

测试场景:
1. 获取包含文本和图片的内容
2. 获取包含文本和表格的内容
3. 使用 maxTextLength 限制文本长度
4. 排除图片和表格，仅获取文本
5. 获取包含详细元数据的内容
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
    options: dict[str, Any],
    wait_seconds: int = 2,
) -> dict[str, Any] | None:
    """执行获取可见内容动作"""
    print("\n📋 获取可见内容...")
    print(f"   选项: {options}")

    await asyncio.sleep(wait_seconds)

    action = OfficeAction(
        category="word",
        action_name="get:visibleContent",
        params={
            "document_uri": document_uri,
            "options": options,
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


def display_content_with_options(data: dict[str, Any], options: dict[str, Any]) -> None:
    """显示内容并验证选项是否生效"""
    print("\n📄 可见内容:")
    text = data.get("text", "")
    metadata = data.get("metadata", {})
    elements = data.get("elements", [])

    # 验证文本
    if options.get("includeText", True):
        print(f"   文本: {text[:100]}{'...' if len(text) > 100 else ''}")
        print(f"   文本总长度: {len(text)}")

        # 验证 maxTextLength：限制单个元素的长度，而非整体文本长度
        max_length = options.get("maxTextLength")
        if max_length:
            exceeded_elements = []
            for i, elem in enumerate(elements):
                if elem.get("type") == "text":
                    content = elem.get("content", {})
                    elem_text = content.get("text", "")
                    if elem_text and len(elem_text) > max_length:
                        exceeded_elements.append(
                            f"元素#{i + 1}: {len(elem_text)} 字符 (限制: {max_length})"
                        )

            if exceeded_elements:
                error_msg = (
                    f"❌ 以下元素超过 maxTextLength 限制:\n"
                    f"   " + "\n   ".join(exceeded_elements)
                )
                print(error_msg)
                assert False, error_msg
            else:
                print(f"   ✅ 所有单个元素长度均 ≤ {max_length} (maxTextLength)")
    else:
        print("   文本: (已排除)")

    # 验证字符数
    print(f"   字符数: {metadata.get('characterCount', 0)}")

    # 验证图片
    if "includeImages" in options:
        image_count = sum(1 for e in elements if e.get("type") == "image")
        if options["includeImages"]:
            # 明确要求包含图片
            print(f"   图片数量: {image_count}")
        else:
            # 明确要求排除图片
            print("   图片: (已排除)")
            if image_count > 0:
                # 找到所有图片元素的详细信息
                image_details = []
                for i, elem in enumerate(elements):
                    if elem.get("type") == "image":
                        content = elem.get("content", {})
                        width = content.get("width", "?")
                        height = content.get("height", "?")
                        alt = content.get("altText", "")
                        detail = f"元素#{i + 1}: width={width}, height={height}"
                        if alt:
                            detail += f", alt={repr(alt[:30])}"
                        image_details.append(detail)

                error_msg = (
                    f"❌ 图片排除失败！设置了 includeImages=False，但返回了 {image_count} 个图片元素。\n"
                    f"   发现的图片元素:\n"
                    f"   " + "\n   ".join(image_details)
                )
                print(error_msg)
                assert False, error_msg
    else:
        # 未设置 includeImages，显示实际数量但不验证
        image_count = sum(1 for e in elements if e.get("type") == "image")
        print(f"   图片数量: {image_count} (未设置选项)")

    # 验证表格
    if "includeTables" in options:
        table_count = sum(1 for e in elements if e.get("type") == "table")
        if options["includeTables"]:
            # 明确要求包含表格
            print(f"   表格数量: {table_count}")
        else:
            # 明确要求排除表格
            print("   表格: (已排除)")
            if table_count > 0:
                # 找到所有表格元素的详细信息
                table_details = []
                for i, elem in enumerate(elements):
                    if elem.get("type") == "table":
                        content = elem.get("content", {})
                        rows = content.get("rows", "?")
                        cols = content.get("columns", "?")
                        detail = f"元素#{i + 1}: rows={rows}, columns={cols}"
                        table_details.append(detail)

                error_msg = (
                    f"❌ 表格排除失败！设置了 includeTables=False，但返回了 {table_count} 个表格元素。\n"
                    f"   发现的表格元素:\n"
                    f"   " + "\n   ".join(table_details)
                )
                print(error_msg)
                assert False, error_msg
    else:
        # 未设置 includeTables，显示实际数量但不验证
        table_count = sum(1 for e in elements if e.get("type") == "table")
        print(f"   表格数量: {table_count} (未设置选项)")

    # 打印详细元数据
    if options.get("detailedMetadata", False):
        print("\n   🔍 详细元数据:")

        # 遍历所有元素，打印详细属性
        for i, elem in enumerate(elements):
            elem_type = elem.get("type")
            content = elem.get("content", {})

            if elem_type == "text":
                # 段落详细属性
                style = content.get("style")
                alignment = content.get("alignment")
                first_line_indent = content.get("firstLineIndent")
                left_indent = content.get("leftIndent")
                right_indent = content.get("rightIndent")
                line_spacing = content.get("lineSpacing")
                space_after = content.get("spaceAfter")
                space_before = content.get("spaceBefore")
                is_list_item = content.get("isListItem")

                details = []
                if style is not None:
                    details.append(f"style={repr(style)}")
                if alignment is not None:
                    details.append(f"alignment={alignment}")
                if first_line_indent is not None:
                    details.append(f"firstLineIndent={first_line_indent}")
                if left_indent is not None:
                    details.append(f"leftIndent={left_indent}")
                if right_indent is not None:
                    details.append(f"rightIndent={right_indent}")
                if line_spacing is not None:
                    details.append(f"lineSpacing={line_spacing}")
                if space_after is not None:
                    details.append(f"spaceAfter={space_after}")
                if space_before is not None:
                    details.append(f"spaceBefore={space_before}")
                if is_list_item is not None:
                    details.append(f"isListItem={is_list_item}")

                if details:
                    text_preview = content.get("text", "")[:30]
                    print(f"      元素#{i + 1} [段落]: {repr(text_preview)}...")
                    for detail in details:
                        print(f"         {detail}")

            elif elem_type == "image":
                # 图片详细属性
                alt_text = content.get("altText")
                hyperlink = content.get("hyperlink")

                details = []
                if alt_text:
                    details.append(f"altText={repr(alt_text)}")
                if hyperlink:
                    details.append(f"hyperlink={repr(hyperlink)}")

                if details:
                    print(f"      元素#{i + 1} [图片]:")
                    for detail in details:
                        print(f"         {detail}")

            elif elem_type == "table":
                # 表格详细属性（cells）
                cells = content.get("cells")
                if cells and len(cells) > 0:
                    print(f"      元素#{i + 1} [表格]: {len(cells)} 行")
                    # 打印前3行的单元格内容
                    for row_idx, row in enumerate(cells[:3]):
                        cell_texts = [cell.get("text", "")[:20] for cell in row]
                        print(f"         行{row_idx}: {cell_texts}")
                    if len(cells) > 3:
                        print(f"         ... (共 {len(cells)} 行)")

            elif elem_type == "other":
                # 内容控件详细属性
                control_type = content.get("controlType")
                title = content.get("title")
                tag = content.get("tag")
                cannot_delete = content.get("cannotDelete")
                cannot_edit = content.get("cannotEdit")
                placeholder_text = content.get("placeholderText")

                details = []
                if control_type is not None:
                    details.append(f"controlType={repr(control_type)}")
                if title:
                    details.append(f"title={repr(title)}")
                if tag:
                    details.append(f"tag={repr(tag)}")
                if cannot_delete is not None:
                    details.append(f"cannotDelete={cannot_delete}")
                if cannot_edit is not None:
                    details.append(f"cannotEdit={cannot_edit}")
                if placeholder_text:
                    details.append(f"placeholderText={repr(placeholder_text)}")

                if details:
                    text_preview = content.get("text", "")[:30]
                    print(f"      元素#{i + 1} [内容控件]: {repr(text_preview)}...")
                    for detail in details:
                        print(f"         {detail}")


async def run_test_template(
    test_name: str,
    test_number: int,
    options: dict[str, Any],
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

            display_content_with_options(data, options)

            print("\n" + "=" * 70)
            print(f"✅ 测试 {test_number} 完成")
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


# ==============================================================================
# 测试函数
# ==============================================================================


async def test_include_text_and_images() -> bool:
    """测试 1: 获取包含文本和图片的内容"""
    return await run_test_template(
        test_name="获取包含文本和图片的内容",
        test_number=1,
        options={"includeText": True, "includeImages": True, "includeTables": False},
    )


async def test_include_text_and_tables() -> bool:
    """测试 2: 获取包含文本和表格的内容"""
    return await run_test_template(
        test_name="获取包含文本和表格的内容",
        test_number=2,
        options={"includeText": True, "includeImages": False, "includeTables": True},
    )


async def test_max_text_length() -> bool:
    """测试 3: 使用 maxTextLength 限制单个元素文本长度"""
    return await run_test_template(
        test_name="使用 maxTextLength 限制单个元素文本长度",
        test_number=3,
        options={"includeText": True, "maxTextLength": 100},
    )


async def test_text_only() -> bool:
    """测试 4: 排除图片和表格，仅获取文本"""
    return await run_test_template(
        test_name="排除图片和表格，仅获取文本",
        test_number=4,
        options={"includeText": True, "includeImages": False, "includeTables": False},
    )


async def test_detailed_metadata() -> bool:
    """测试 5: 获取包含详细元数据的内容"""
    return await run_test_template(
        test_name="获取包含详细元数据的内容",
        test_number=5,
        options={"includeText": True, "detailedMetadata": True},
    )


async def run_all_tests() -> bool:
    """运行所有选项测试"""
    print("\n🚀 运行所有选项获取可见内容测试...\n")
    results = []
    results.append(await test_include_text_and_images())
    await asyncio.sleep(2)
    results.append(await test_include_text_and_tables())
    await asyncio.sleep(2)
    results.append(await test_max_text_length())
    await asyncio.sleep(2)
    results.append(await test_text_only())
    await asyncio.sleep(2)
    results.append(await test_detailed_metadata())
    success = all(results)

    print("\n" + "=" * 70)
    print(f"📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    print("=" * 70)
    return success


# ==============================================================================
# 主程序入口
# ==============================================================================

TEST_MAPPING = {
    "1": test_include_text_and_images,
    "2": test_include_text_and_tables,
    "3": test_max_text_length,
    "4": test_text_only,
    "5": test_detailed_metadata,
    "all": run_all_tests,
}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Get Visible Content E2E Tests - Options")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "5", "all"],
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
