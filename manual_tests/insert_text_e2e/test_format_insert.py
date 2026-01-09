"""
Formatted Text Insert Tests

测试带格式（format 参数）的文本插入功能。

测试场景:
1. 粗体文本插入
2. 斜体文本插入
3. 字体大小设置
4. 字体名称设置
5. 颜色设置
6. 组合格式（粗体+斜体+大小+颜色）
"""

import asyncio
import sys

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


async def test_bold_text_insert():
    """
    测试 1: 插入粗体文本
    format={bold: true}
    """
    print("\n" + "=" * 70)
    print("🧪 测试 1: 插入粗体文本")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        await workspace.start()
        print("✅ Workspace 启动成功")

        # 等待连接
        print("\n⏳ 等待 Word Add-In 连接...")
        connected = await workspace.wait_for_addin_connection(timeout=30.0)
        if not connected:
            print("❌ 超时：未检测到 Add-In 连接")
            return False

        # 获取文档
        documents = workspace.get_connected_documents()
        if not documents:
            print("❌ 未找到已连接文档")
            return False

        test_document_uri = documents[0]
        print(f"✅ 使用文档: {test_document_uri}")

        # 执行插入动作
        print("\n📝 插入粗体文本: '这是粗体文本'")
        print("   格式: bold=true")
        print("   提示: 请将光标放在要插入文本的位置")

        await asyncio.sleep(3)

        action = OfficeAction(
            category="word",
            action_name="insert:text",
            params={
                "document_uri": test_document_uri,
                "text": "这是粗体文本",
                "location": "Cursor",
                "format": {"bold": True},
            },
        )

        result = await workspace.execute(action)

        # 验证结果
        print("\n📊 验证结果:")
        if result.success:
            print("✅ 插入成功")
            print(f"   返回数据: {result.data}")
            print("\n   请检查 Word 文档，确认文本为粗体")
        else:
            print(f"❌ 插入失败: {result.error}")
            return False

        print("\n" + "=" * 70)
        print("✅ 测试 1 完成")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        await workspace.stop()
        print("✅ Workspace 已停止")


async def test_italic_text_insert():
    """
    测试 2: 插入斜体文本
    format={italic: true}
    """
    print("\n" + "=" * 70)
    print("🧪 测试 2: 插入斜体文本")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        await workspace.start()
        print("✅ Workspace 启动成功")

        # 等待连接
        print("\n⏳ 等待 Word Add-In 连接...")
        connected = await workspace.wait_for_addin_connection(timeout=30.0)
        if not connected:
            print("❌ 超时：未检测到 Add-In 连接")
            return False

        # 获取文档
        documents = workspace.get_connected_documents()
        if not documents:
            print("❌ 未找到已连接文档")
            return False

        test_document_uri = documents[0]
        print(f"✅ 使用文档: {test_document_uri}")

        # 执行插入动作
        print("\n📝 插入斜体文本: '这是斜体文本'")
        print("   格式: italic=true")
        print("   提示: 请将光标放在要插入文本的位置")

        await asyncio.sleep(3)

        action = OfficeAction(
            category="word",
            action_name="insert:text",
            params={
                "document_uri": test_document_uri,
                "text": "这是斜体文本",
                "location": "Cursor",
                "format": {"italic": True},
            },
        )

        result = await workspace.execute(action)

        # 验证结果
        print("\n📊 验证结果:")
        if result.success:
            print("✅ 插入成功")
            print(f"   返回数据: {result.data}")
            print("\n   请检查 Word 文档，确认文本为斜体")
        else:
            print(f"❌ 插入失败: {result.error}")
            return False

        print("\n" + "=" * 70)
        print("✅ 测试 2 完成")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        await workspace.stop()
        print("✅ Workspace 已停止")


async def test_font_size_insert():
    """
    测试 3: 插入不同字体大小的文本
    format={fontSize: 20}
    """
    print("\n" + "=" * 70)
    print("🧪 测试 3: 插入不同字体大小的文本")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        await workspace.start()
        print("✅ Workspace 启动成功")

        # 等待连接
        print("\n⏳ 等待 Word Add-In 连接...")
        connected = await workspace.wait_for_addin_connection(timeout=30.0)
        if not connected:
            print("❌ 超时：未检测到 Add-In 连接")
            return False

        # 获取文档
        documents = workspace.get_connected_documents()
        if not documents:
            print("❌ 未找到已连接文档")
            return False

        test_document_uri = documents[0]
        print(f"✅ 使用文档: {test_document_uri}")

        # 测试三个不同的字体大小
        test_cases = [
            (12, "小号文本 (12pt)"),
            (16, "中号文本 (16pt)"),
            (24, "大号文本 (24pt)"),
        ]

        print("\n📝 将插入 3 个不同字体大小的文本")
        print("   提示: 请将光标放在要插入文本的位置")

        await asyncio.sleep(3)

        results = []
        for font_size, text in test_cases:
            print(f"\n--- 插入 {text} ---")

            action = OfficeAction(
                category="word",
                action_name="insert:text",
                params={
                    "document_uri": test_document_uri,
                    "text": f"{text}\n",
                    "location": "Cursor",
                    "format": {"fontSize": font_size},
                },
            )

            result = await workspace.execute(action)
            results.append(result.success)

            if result.success:
                print(f"✅ 插入成功 (fontSize={font_size})")
            else:
                print(f"❌ 插入失败: {result.error}")

            await asyncio.sleep(1)

        # 验证结果
        print("\n📊 验证结果:")
        success_count = sum(results)
        print(f"   成功: {success_count}/{len(results)}")

        if all(results):
            print("\n   ✅ 所有插入都成功！")
            print("   请检查 Word 文档，确认字体大小递增效果")
        else:
            print("\n   ⚠️  部分插入失败")
            return False

        print("\n" + "=" * 70)
        print("✅ 测试 3 完成")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        await workspace.stop()
        print("✅ Workspace 已停止")


async def test_font_name_insert():
    """
    测试 4: 插入不同字体的文本
    format={fontName: "Arial"}
    """
    print("\n" + "=" * 70)
    print("🧪 测试 4: 插入不同字体的文本")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        await workspace.start()
        print("✅ Workspace 启动成功")

        # 等待连接
        print("\n⏳ 等待 Word Add-In 连接...")
        connected = await workspace.wait_for_addin_connection(timeout=30.0)
        if not connected:
            print("❌ 超时：未检测到 Add-In 连接")
            return False

        # 获取文档
        documents = workspace.get_connected_documents()
        if not documents:
            print("❌ 未找到已连接文档")
            return False

        test_document_uri = documents[0]
        print(f"✅ 使用文档: {test_document_uri}")

        # 测试不同的字体
        test_cases = [
            ("Arial", "Arial 字体"),
            ("Times New Roman", "Times New Roman 字体"),
            ("Courier New", "Courier New 字体"),
        ]

        print("\n📝 将插入 3 个不同字体的文本")
        print("   提示: 请将光标放在要插入文本的位置")

        await asyncio.sleep(3)

        results = []
        for font_name, text in test_cases:
            print(f"\n--- 插入 {text} ---")

            action = OfficeAction(
                category="word",
                action_name="insert:text",
                params={
                    "document_uri": test_document_uri,
                    "text": f"{text}\n",
                    "location": "Cursor",
                    "format": {"fontName": font_name},
                },
            )

            result = await workspace.execute(action)
            results.append(result.success)

            if result.success:
                print(f"✅ 插入成功 (fontName={font_name})")
            else:
                print(f"❌ 插入失败: {result.error}")

            await asyncio.sleep(1)

        # 验证结果
        print("\n📊 验证结果:")
        success_count = sum(results)
        print(f"   成功: {success_count}/{len(results)}")

        if all(results):
            print("\n   ✅ 所有插入都成功！")
            print("   请检查 Word 文档，确认不同字体效果")
        else:
            print("\n   ⚠️  部分插入失败")
            return False

        print("\n" + "=" * 70)
        print("✅ 测试 4 完成")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        await workspace.stop()
        print("✅ Workspace 已停止")


async def test_color_insert():
    """
    测试 5: 插入不同颜色的文本
    format={color: "#FF0000"}
    """
    print("\n" + "=" * 70)
    print("🧪 测试 5: 插入不同颜色的文本")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        await workspace.start()
        print("✅ Workspace 启动成功")

        # 等待连接
        print("\n⏳ 等待 Word Add-In 连接...")
        connected = await workspace.wait_for_addin_connection(timeout=30.0)
        if not connected:
            print("❌ 超时：未检测到 Add-In 连接")
            return False

        # 获取文档
        documents = workspace.get_connected_documents()
        if not documents:
            print("❌ 未找到已连接文档")
            return False

        test_document_uri = documents[0]
        print(f"✅ 使用文档: {test_document_uri}")

        # 测试不同的颜色
        test_cases = [
            ("#FF0000", "红色文本"),
            ("#00FF00", "绿色文本"),
            ("#0000FF", "蓝色文本"),
        ]

        print("\n📝 将插入 3 个不同颜色的文本")
        print("   提示: 请将光标放在要插入文本的位置")

        await asyncio.sleep(3)

        results = []
        for color, text in test_cases:
            print(f"\n--- 插入 {text} ---")

            action = OfficeAction(
                category="word",
                action_name="insert:text",
                params={
                    "document_uri": test_document_uri,
                    "text": f"{text}\n",
                    "location": "Cursor",
                    "format": {"color": color},
                },
            )

            result = await workspace.execute(action)
            results.append(result.success)

            if result.success:
                print(f"✅ 插入成功 (color={color})")
            else:
                print(f"❌ 插入失败: {result.error}")

            await asyncio.sleep(1)

        # 验证结果
        print("\n📊 验证结果:")
        success_count = sum(results)
        print(f"   成功: {success_count}/{len(results)}")

        if all(results):
            print("\n   ✅ 所有插入都成功！")
            print("   请检查 Word 文档，确认不同颜色效果")
        else:
            print("\n   ⚠️  部分插入失败")
            return False

        print("\n" + "=" * 70)
        print("✅ 测试 5 完成")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        await workspace.stop()
        print("✅ Workspace 已停止")


async def test_combined_format_insert():
    """
    测试 6: 插入带组合格式的文本
    format={bold: true, italic: true, fontSize: 18, fontName: "Arial", color: "#FF0000"}
    """
    print("\n" + "=" * 70)
    print("🧪 测试 6: 插入带组合格式的文本")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        await workspace.start()
        print("✅ Workspace 启动成功")

        # 等待连接
        print("\n⏳ 等待 Word Add-In 连接...")
        connected = await workspace.wait_for_addin_connection(timeout=30.0)
        if not connected:
            print("❌ 超时：未检测到 Add-In 连接")
            return False

        # 获取文档
        documents = workspace.get_connected_documents()
        if not documents:
            print("❌ 未找到已连接文档")
            return False

        test_document_uri = documents[0]
        print(f"✅ 使用文档: {test_document_uri}")

        # 执行插入动作
        print("\n📝 插入带组合格式的文本: '组合格式文本'")
        print("   格式:")
        print("     - bold: true")
        print("     - italic: true")
        print("     - fontSize: 18")
        print("     - fontName: Arial")
        print("     - color: #FF0000 (红色)")
        print("\n   提示: 请将光标放在要插入文本的位置")

        await asyncio.sleep(3)

        action = OfficeAction(
            category="word",
            action_name="insert:text",
            params={
                "document_uri": test_document_uri,
                "text": "组合格式文本",
                "location": "Cursor",
                "format": {
                    "bold": True,
                    "italic": True,
                    "fontSize": 18,
                    "fontName": "Arial",
                    "color": "#FF0000",
                },
            },
        )

        result = await workspace.execute(action)

        # 验证结果
        print("\n📊 验证结果:")
        if result.success:
            print("✅ 插入成功")
            print(f"   返回数据: {result.data}")
            print("\n   请检查 Word 文档，确认组合格式效果")
            print("   (粗体、斜体、18号、Arial、红色)")
        else:
            print(f"❌ 插入失败: {result.error}")
            return False

        print("\n" + "=" * 70)
        print("✅ 测试 6 完成")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        await workspace.stop()
        print("✅ Workspace 已停止")


async def run_all_tests():
    """运行所有格式插入测试"""
    print("\n🚀 运行所有格式插入测试...\n")
    results = []
    results.append(await test_bold_text_insert())
    await asyncio.sleep(2)
    results.append(await test_italic_text_insert())
    await asyncio.sleep(2)
    results.append(await test_font_size_insert())
    await asyncio.sleep(2)
    results.append(await test_font_name_insert())
    await asyncio.sleep(2)
    results.append(await test_color_insert())
    await asyncio.sleep(2)
    results.append(await test_combined_format_insert())
    success = all(results)

    print("\n" + "=" * 70)
    print(f"📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    print("=" * 70)
    return success


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Formatted Text Insert E2E Tests")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "5", "6", "all"],
        default="1",
        help="Test to run: 1=bold, 2=italic, 3=fontSize, 4=fontName, 5=color, 6=combined, all=all tests",
    )

    args = parser.parse_args()

    try:
        if args.test == "1":
            success = asyncio.run(test_bold_text_insert())
        elif args.test == "2":
            success = asyncio.run(test_italic_text_insert())
        elif args.test == "3":
            success = asyncio.run(test_font_size_insert())
        elif args.test == "4":
            success = asyncio.run(test_font_name_insert())
        elif args.test == "5":
            success = asyncio.run(test_color_insert())
        elif args.test == "6":
            success = asyncio.run(test_combined_format_insert())
        else:  # all
            success = asyncio.run(run_all_tests())

        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
