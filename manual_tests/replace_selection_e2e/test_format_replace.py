"""
Format Replace Test

测试带格式的文本替换功能。

测试场景:
1. 替换为粗体文本
2. 替换为斜体文本
3. 替换为带字体格式的文本
4. 替换为带颜色和下划线的文本

Usage:
    # 运行单个测试
    uv run python manual_tests/replace_selection_e2e/test_format_replace.py --test 1

    # 运行全部测试
    uv run python manual_tests/replace_selection_e2e/test_format_replace.py --test all
"""

import asyncio
import sys
from contextlib import asynccontextmanager

# Add project root to path
sys.path.insert(0, "/Users/jqq/PycharmProjects/office4ai")

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


# ==============================================================================
# 辅助函数和上下文管理器
# ==============================================================================


@asynccontextmanager
async def workspace_context(host: str = "127.0.0.1", port: int = 3000):
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
    return documents[0]


async def replace_selection(
    workspace: OfficeWorkspace,
    document_uri: str,
    content: dict,
    wait_seconds: int = 3,
) -> bool:
    """
    执行选择替换动作

    Args:
        workspace: Workspace 实例
        document_uri: 目标文档 URI
        content: 替换内容
        wait_seconds: 执行前等待秒数

    Returns:
        bool: 是否成功
    """
    print(f"\n📝 替换选择: {content}")
    print(f"   等待 {wait_seconds} 秒...")

    # Wait for user to select text
    await asyncio.sleep(wait_seconds)

    # Create action
    action = OfficeAction(
        category="word",
        action_name="replace:selection",
        params={
            "document_uri": document_uri,
            "content": content,
        },
    )

    # Execute action
    observation = await workspace.execute(action)

    return observation.success if observation else False


# ==============================================================================
# 测试用例
# ==============================================================================


async def test_1_replace_with_bold_text() -> None:
    """
    测试场景 1: 替换为粗体文本

    步骤:
        1. 启动 Workspace Socket.IO 服务器
        2. 等待 Word Add-In 连接
        3. 在 Word 中选中一些文本
        4. 发送 word:replace:selection 请求，替换为粗体文本
        5. 验证替换成功并检查格式

    预期结果:
        - 选中的文本被替换为粗体 "Bold Text"
        - 返回 replaced=True
        - 文本在 Word 中显示为粗体
    """
    print("\n" + "=" * 60)
    print("测试 1: 替换为粗体文本")
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

        content = {"text": "Bold Text", "format": {"bold": True}}
        success = await replace_selection(workspace, document_uri, content)

        print("\n🔍 验证:")
        if success:
            print("   ✅ 测试通过: 粗体文本替换成功")
            print("   👀 请检查 Word 中的文本是否为粗体")
        else:
            print("   ❌ 测试失败: 替换失败")


async def test_2_replace_with_italic_text() -> None:
    """
    测试场景 2: 替换为斜体文本

    步骤:
        1. 在 Word 中选中一些文本
        2. 发送 word:replace:selection 请求，替换为斜体文本

    预期结果:
        - 选中的文本被替换为斜体文本
        - 文本在 Word 中显示为斜体
    """
    print("\n" + "=" * 60)
    print("测试 2: 替换为斜体文本")
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

        content = {"text": "Italic Text", "format": {"italic": True}}
        success = await replace_selection(workspace, document_uri, content)

        print("\n🔍 验证:")
        if success:
            print("   ✅ 测试通过: 斜体文本替换成功")
            print("   👀 请检查 Word 中的文本是否为斜体")
        else:
            print("   ❌ 测试失败")


async def test_3_replace_with_font_format() -> None:
    """
    测试场景 3: 替换为带字体格式的文本

    步骤:
        1. 在 Word 中选中一些文本
        2. 发送 word:replace:selection 请求，替换为带字体名称和大小的文本

    预期结果:
        - 选中的文本被替换
        - 字体名称为 Arial，大小为 16pt
    """
    print("\n" + "=" * 60)
    print("测试 3: 替换为带字体格式的文本")
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

        content = {
            "text": "Formatted Text",
            "format": {"fontName": "Arial", "fontSize": 16, "bold": True, "italic": False},
        }
        success = await replace_selection(workspace, document_uri, content)

        print("\n🔍 验证:")
        if success:
            print("   ✅ 测试通过: 字体格式替换成功")
            print("   👀 请检查 Word 中的字体是否为 Arial 16pt 粗体")
        else:
            print("   ❌ 测试失败")


async def test_4_replace_with_color_and_underline() -> None:
    """
    测试场景 4: 替换为带颜色和下划线的文本

    步骤:
        1. 在 Word 中选中一些文本
        2. 发送 word:replace:selection 请求，替换为带颜色和下划线的文本

    预期结果:
        - 选中的文本被替换
        - 文本颜色为红色 (#FF0000)
        - 文本带下划线
    """
    print("\n" + "=" * 60)
    print("测试 4: 替换为带颜色和下划线的文本")
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

        content = {
            "text": "Colorful Underlined Text",
            "format": {"color": "#FF0000", "underline": True, "bold": True},
        }
        success = await replace_selection(workspace, document_uri, content)

        print("\n🔍 验证:")
        if success:
            print("   ✅ 测试通过: 颜色和下划线替换成功")
            print("   👀 请检查 Word 中的文本是否为红色、带下划线、粗体")
        else:
            print("   ❌ 测试失败")


# ==============================================================================
# 测试运行器
# ==============================================================================


async def main() -> None:
    """主函数：运行指定的测试"""
    import argparse

    parser = argparse.ArgumentParser(description="Format Replace E2E Tests")
    parser.add_argument(
        "--test",
        type=str,
        required=True,
        help="Test number to run (1-4) or 'all'",
    )
    args = parser.parse_args()

    tests = {
        "1": ("替换为粗体文本", test_1_replace_with_bold_text),
        "2": ("替换为斜体文本", test_2_replace_with_italic_text),
        "3": ("替换为带字体格式的文本", test_3_replace_with_font_format),
        "4": ("替换为带颜色和下划线的文本", test_4_replace_with_color_and_underline),
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
