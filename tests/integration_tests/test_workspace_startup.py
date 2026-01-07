"""
Workspace Startup Test

测试 Office Workspace 能够成功启动 Socket.IO 服务器。
"""

import asyncio
import sys
from pathlib import Path

# 添加项目根目录到 Python 路径
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from office4ai.environment.workspace.office_workspace import OfficeWorkspace


async def test_workspace_startup():
    """
    测试 Workspace 启动和基本功能

    这个测试会：
    1. 启动 Workspace Socket.IO 服务器
    2. 等待用户按 Enter 停止
    3. 清理资源
    """
    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    try:
        # 启动 Workspace
        print("\n🚀 Starting Office Workspace...")
        await workspace.start()

        print("\n✅ Workspace is running!")
        print("   - Health check: http://127.0.0.1:3000/health")
        print("   - Socket.IO Server: ws://127.0.0.1:3000/socket.io/")
        print("   - Word namespace: /word")
        print("\n💡 Next steps:")
        print("   1. 在 Word 中加载 office-editor4ai Add-In")
        print("   2. Add-In 会自动连接到 Workspace")
        print("   3. 按 Ctrl+C 或 Enter 停止服务器")

        # 等待 Add-In 连接
        print("\n⏳ Waiting for Add-In connection (300s timeout)...")
        connected = await workspace.wait_for_addin_connection(timeout=300.0)

        if connected:
            docs = workspace.get_connected_documents()
            print(f"✅ Connected documents: {docs}")

        # 保持运行直到用户中断
        print("\n⌨️  Press Enter or Ctrl+C to stop...")
        await asyncio.sleep(3600)  # 保持运行 1 小时

    except KeyboardInterrupt:
        print("\n\n⏸️  Received interrupt signal")
    finally:
        # 停止 Workspace
        print("\n🛑 Stopping Office Workspace...")
        await workspace.stop()
        print("✅ Workspace stopped")


if __name__ == "__main__":
    try:
        asyncio.run(test_workspace_startup())
    except KeyboardInterrupt:
        print("\n✅ Test completed")
