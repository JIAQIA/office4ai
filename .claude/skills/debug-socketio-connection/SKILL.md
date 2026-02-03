---
name: debug-socketio-connection
description: 调试 Socket.IO 连接问题，特别是 "Document not connected" 错误。当 Add-In 已连接但执行命令时报文档未连接错误时使用此 skill。
---

# Debug Socket.IO Connection Issues

## 问题模式

当看到如下错误时，说明 Add-In 与 Server 之间存在 URI 匹配问题：

```
✅ Add-In 已连接
📝 执行: 获取文档统计...
❌ 获取失败: Document not connected: file:///path/to/doc.docx
```

**核心矛盾**：Add-In 确实已连接（日志显示 "Client registered"），但执行命令时却报 "Document not connected"。

## 根因分析

这类问题的根本原因是 **URI 不匹配**。ConnectionManager 使用 `document_uri` 作为 key 来路由请求，如果注册时和查询时的 URI 不一致，就会找不到连接。

### 常见的 URI 不匹配场景

| 场景 | 注册时 URI | 查询时 URI | 原因 |
|------|-----------|-----------|------|
| URL 编码 | `file:///%2Fvar%2Ffolders%2F...` | `file:///var/folders/...` | Add-In 对路径进行了 URL 编码 |
| 符号链接 | `file:///var/folders/...` | `file:///private/var/folders/...` | macOS `/var` → `/private/var` |
| 路径格式 | `file:////server/share/...` | `file:///server/share/...` | UNC 路径的斜杠数量 |

## 调试步骤

### Step 1: 对比日志中的 URI

查看 Add-In 连接时的日志，找到 **注册时的 URI**：

```
Client registered: word-client-xxx (socket_id) for <注册URI> on /word
```

对比 **查询时的 URI**（错误信息中）：

```
Document not connected: <查询URI>
```

### Step 2: 检查 URI 规范化逻辑

URI 规范化函数位于 [connection_manager.py](../../office4ai/environment/workspace/socketio/services/connection_manager.py) 的 `normalize_document_uri()` 函数。

该函数处理：
1. URL 解码（`%2F` → `/`）
2. 符号链接解析（`/var` → `/private/var`）
3. 路径规范化

### Step 3: 确认所有入口点都做了规范化

以下方法必须在处理前调用 `normalize_document_uri()`：

- `register_client()` - 注册时规范化
- `get_socket_by_document()` - 查询时规范化
- `get_clients_by_document()` - 查询时规范化
- `is_document_active()` - 查询时规范化

### Step 4: 添加测试用例验证

参考 [test_connection_manager.py](../../tests/unit_tests/office4ai/environment/workspace/socketio/services/test_connection_manager.py) 中的 `TestNormalizeDocumentUri` 类，确保：

```python
def test_lookup_with_different_uri_formats(self, connection_manager):
    """注册和查询使用不同格式的 URI 应该能匹配"""
    # 用 URL 编码格式注册
    connection_manager.register_client(
        "socket1", "client1", "file:///%2Ftmp%2Ftest.docx", "/word"
    )
    # 用解码格式查询应该能找到
    socket_id = connection_manager.get_socket_by_document("file:///tmp/test.docx")
    assert socket_id == "socket1"
```

## 预防措施

### 测试文件位置

避免使用系统临时目录（如 `/var/folders/`），因为可能涉及符号链接。

参考 [e2e_base.py](../../manual_tests/e2e_base.py) 的做法，使用项目内目录：

```python
# 使用项目内目录，避免符号链接问题
TEMP_ROOT = Path(__file__).parent / ".test_working"
```

### URI 生成函数

确保 URI 生成函数使用 `os.path.realpath()` 解析符号链接：

```python
def path_to_file_uri(path: Path) -> str:
    abs_path = path.resolve()  # 解析符号链接
    encoded = quote(str(abs_path), safe="/")
    return f"file://{encoded}"
```

## 注意事项

### 清理顺序很重要

测试结束后的清理顺序必须正确，否则会导致长时间等待：

```
❌ 错误顺序（等待 ~2 分钟）：
1. workspace.stop()  ← 等待 Socket.IO 连接超时
2. close_document()  ← 关闭文档

✅ 正确顺序（~1 秒完成）：
1. close_document()  ← 先关闭文档，Add-In 自动断开
2. workspace.stop()  ← 没有活跃连接，立即完成
```

**原因**：Socket.IO 服务器在关闭时会等待所有活跃连接断开或超时。如果先关闭服务器，而文档还开着（Add-In 还连接着），就会等待 ping 超时（默认 60-120 秒）。

参考 [e2e_base.py](../../manual_tests/e2e_base.py) 中 `run_with_workspace` 的 finally 块实现。

## 相关文件

- 连接管理器实现：[connection_manager.py](../../office4ai/environment/workspace/socketio/services/connection_manager.py)
- 连接管理器测试：[test_connection_manager.py](../../tests/unit_tests/office4ai/environment/workspace/socketio/services/test_connection_manager.py)
- E2E 测试基础设施：[e2e_base.py](../../manual_tests/e2e_base.py)
