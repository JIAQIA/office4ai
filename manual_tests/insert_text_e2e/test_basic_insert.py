"""
Basic Insert Text Test

测试基本的文本插入功能（使用默认参数）。

测试场景:
1. 插入纯文本（使用默认 location="Cursor"）
2. 插入多行文本
3. 插入特殊字符
4. 插入长文本
"""

import asyncio
import sys
from contextlib import asynccontextmanager

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


async def insert_text(
    workspace: OfficeWorkspace,
    document_uri: str,
    text: str,
    wait_seconds: int = 3,
) -> bool:
    """
    执行文本插入动作

    Args:
        workspace: Workspace 实例
        document_uri: 目标文档 URI
        text: 要插入的文本
        wait_seconds: 执行前等待秒数

    Returns:
        bool: 是否成功
    """
    print(f"\n📝 插入文本: '{text[:50]}{'...' if len(text) > 50 else ''}'")
    print(f"   长度: {len(text)} 字符")
    print("   提示: 请将光标放在要插入文本的位置")

    await asyncio.sleep(wait_seconds)

    action = OfficeAction(
        category="word",
        action_name="insert:text",
        params={
            "document_uri": document_uri,
            "text": text,
        },
    )

    result = await workspace.execute(action)

    # 验证结果
    print("\n📊 验证结果:")
    if result.success:
        print("✅ 插入成功")
        print(f"   返回数据: {result.data}")
        return True
    else:
        print(f"❌ 插入失败: {result.error}")
        return False


async def run_test_template(
    test_name: str,
    test_number: int,
    text: str,
    wait_seconds: int = 3,
) -> bool:
    """
    测试执行模板：封装通用的测试流程

    Args:
        test_name: 测试名称
        test_number: 测试编号
        text: 要插入的文本
        wait_seconds: 执行前等待秒数

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

            print(f"✅ 使用文档: {document_uri}")

            # 执行插入
            if not await insert_text(workspace, document_uri, text, wait_seconds):
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


async def test_simple_text_insert():
    """测试 1: 简单文本插入"""
    return await run_test_template(
        test_name="简单文本插入",
        test_number=1,
        text="Hello World",
    )


async def test_multiline_text_insert():
    """测试 2: 多行文本插入"""
    multiline_text = """第一行文本
第二行文本
第三行文本"""
    return await run_test_template(
        test_name="多行文本插入",
        test_number=2,
        text=multiline_text,
    )


async def test_special_characters_insert():
    """测试 3: 特殊字符插入"""
    special_text = "特殊字符测试: @#$%^&*()_+-=[]{}|;':\",./<>?~`"
    return await run_test_template(
        test_name="特殊字符插入",
        test_number=3,
        text=special_text,
    )


async def test_long_text_insert():
    """测试 4: 长文本插入"""
    long_text = """
这是一段较长的文本，用于测试 Word Add-In 处理较长内容的能力。
这段文本包含了多个句子，每个句子都测试不同的字符和标点符号。
在插入这段文本后，我们应该验证：

1. 文本是否完整插入
2. 格式是否保持正确
3. 是否有乱码或丢失字符

此外，我们还需要测试性能，确保插入长文本不会导致系统卡顿或超时。
这个测试对于确保用户体验非常重要，因为在实际使用中，用户可能会插入大段文本。
""".strip()
    return await run_test_template(
        test_name="长文本插入",
        test_number=4,
        text=long_text,
    )


async def run_all_tests():
    """运行所有基本插入测试"""
    print("\n🚀 运行所有基本插入测试...\n")
    results = []
    results.append(await test_simple_text_insert())
    await asyncio.sleep(2)
    results.append(await test_multiline_text_insert())
    await asyncio.sleep(2)
    results.append(await test_special_characters_insert())
    await asyncio.sleep(2)
    results.append(await test_long_text_insert())
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
    "1": test_simple_text_insert,
    "2": test_multiline_text_insert,
    "3": test_special_characters_insert,
    "4": test_long_text_insert,
    "all": run_all_tests,
}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Basic Insert Text E2E Tests")
    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "all"],
        default="1",
        help="Test to run: 1=simple, 2=multiline, 3=special, 4=long, all=all tests",
    )

    args = parser.parse_args()

    try:
        test_func = TEST_MAPPING[args.test]
        success = asyncio.run(test_func())
        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
