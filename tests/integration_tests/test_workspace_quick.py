"""
Workspace Quick Test

快速测试 Workspace 能够启动和停止。
"""

import asyncio

from office4ai.environment.workspace.office_workspace import OfficeWorkspace


async def test_workspace_quick():
    """快速测试：启动 → 等待 2 秒 → 停止"""
    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)

    print("\n🚀 Starting Office Workspace...")
    await workspace.start()

    print("✅ Workspace started successfully!")
    print(f"   Running: {workspace.is_running}")
    print(f"   Connections: {workspace.sio_server is not None}")

    # 等待 2 秒
    print("\n⏳ Waiting 2 seconds...")
    await asyncio.sleep(2)

    # 停止
    print("\n🛑 Stopping Workspace...")
    await workspace.stop()

    print("✅ Test passed!")


if __name__ == "__main__":
    asyncio.run(test_workspace_quick())
