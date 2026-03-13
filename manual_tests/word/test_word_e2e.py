"""
Word End-to-End Integration Test

测试 Workspace 与 Word Add-In 的完整通信流程。

测试流程:
1. 启动 Office Workspace Socket.IO 服务器
2. 等待 Word Add-In 连接
3. 调用 word:get:selectedContent 获取选中内容
4. 验证返回结果
5. 清理资源

要求:
- Workspace 服务器运行在 http://127.0.0.1:3000
- Word Add-In 已加载并连接到服务器
- Word 文档中有选中的文本
"""

import asyncio
import sys

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


async def test_word_e2e():
    """
    端到端测试：Workspace → Socket.IO → Word Add-In → 响应

    测试场景:
    1. 启动 Workspace
    2. 等待 Add-In 连接 (30s timeout)
    3. 获取已连接文档列表
    4. 调用 word:get:selectedContent
    5. 验证返回结果
    """

    print("\n" + "=" * 70)
    print("🧪 Word End-to-End Integration Test")
    print("=" * 70)

    # Step 1: 启动 Workspace
    print("\n[1/5] 启动 Office Workspace...")
    # 使用默认配置（use_https=True）以匹配客户端的默认连接地址 https://127.0.0.1:4443
    # Use default config (use_https=True) to match client's default connection URL https://127.0.0.1:4443
    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        await workspace.start()
        print("✅ Workspace 启动成功")
        print(f"   运行状态: {workspace.is_running}")

        # Step 2: 等待 Add-In 连接
        print("\n[2/5] 等待 Word Add-In 连接...")
        print("   提示: 请在 Word 中加载 office-editor4ai Add-In")
        print("   超时时间: 30 秒")

        connected = await workspace.wait_for_addin_connection(timeout=30.0)

        if not connected:
            print("❌ 超时：未检测到 Add-In 连接")
            print("   请检查:")
            print("   1. Word Add-In 是否已加载")
            print("   2. Add-In 是否能访问 http://127.0.0.1:3000")
            print("   3. 浏览器控制台是否有错误")
            return False

        print("✅ Add-In 已连接")

        # Step 3: 获取已连接文档列表
        print("\n[3/5] 获取已连接文档...")
        documents = workspace.get_connected_documents()

        if not documents:
            print("⚠️  未找到已连接文档")
            return False

        print(f"✅ 找到 {len(documents)} 个已连接文档:")
        for i, doc_uri in enumerate(documents, 1):
            print(f"   {i}. {doc_uri}")

        # 使用第一个文档进行测试
        test_document_uri = documents[0]
        print(f"\n   使用文档进行测试: {test_document_uri}")

        # Step 4: 调用 word:get:selectedContent
        print("\n[4/5] 调用 word:get:selectedContent...")
        print("   提示: 请在 Word 文档中选中一些文本")

        # 等待用户选中文本
        print("   等待 5 秒...")
        await asyncio.sleep(5)

        # 创建动作
        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={
                "document_uri": test_document_uri,
            },
        )

        print(f"   发送动作: {action.category}:{action.action_name}")

        # 执行动作
        result = await workspace.execute(action)

        # Step 5: 验证结果
        print("\n[5/5] 验证结果...")

        if not result.success:
            print(f"❌ 动作执行失败: {result.error}")
            return False

        print("✅ 动作执行成功")
        print(f"   返回数据: {result.data}")

        # 检查返回的内容
        # result.data 直接包含 ContentInfo 对象: {text, elements, metadata}
        if isinstance(result.data, dict):
            text = result.data.get("text", "")
            elements = result.data.get("elements", [])
            metadata = result.data.get("metadata", {})

            print("\n   📝 选中文本内容:")
            print(f"   '{text}'")

            print(f"\n   📊 元素数量: {len(elements)}")
            if elements:
                print(f"   第一个元素类型: {elements[0].get('type')}")

            print("\n   📈 统计信息:")
            print(f"   - 字符数: {metadata.get('characterCount', 0)}")
            print(f"   - 段落数: {metadata.get('paragraphCount', 0)}")
            print(f"   - 表格数: {metadata.get('tableCount', 0)}")
            print(f"   - 图片数: {metadata.get('imageCount', 0)}")

            if text:
                print(f"\n   ✅ 成功获取选中文本 (长度: {len(text)})")
            else:
                print("\n   ⚠️  选中文本为空 (可能未选中内容)")
        else:
            print(f"\n   ⚠️  返回数据格式异常: {type(result.data)}")

        print("\n" + "=" * 70)
        print("✅ 端到端测试完成")
        print("=" * 70)

        return True

    except Exception as e:
        print(f"\n❌ 测试过程中发生错误: {e}")
        import traceback

        traceback.print_exc()
        return False

    finally:
        # 清理资源
        print("\n🧹 清理资源...")
        await workspace.stop()
        print("✅ Workspace 已停止")


async def test_workspace_health():
    """
    快速健康检查：验证 Workspace 能够正常启动和响应
    """
    print("\n" + "=" * 70)
    print("🏥 Workspace Health Check")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        print("\n启动 Workspace...")
        await workspace.start()

        print("✅ Workspace 运行正常")
        print(f"   运行状态: {workspace.is_running}")
        print("   健康检查: http://127.0.0.1:3000/health")

        # 等待 2 秒
        await asyncio.sleep(2)

        return True

    except Exception as e:
        print(f"❌ 健康检查失败: {e}")
        return False

    finally:
        await workspace.stop()
        print("✅ Workspace 已停止")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Word End-to-End Integration Tests")
    parser.add_argument(
        "--mode",
        choices=["health", "e2e"],
        default="health",
        help="Test mode: health (quick check) or e2e (full integration test)",
    )

    args = parser.parse_args()

    try:
        if args.mode == "health":
            success = asyncio.run(test_workspace_health())
            sys.exit(0 if success else 1)
        else:
            success = asyncio.run(test_word_e2e())
            sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
