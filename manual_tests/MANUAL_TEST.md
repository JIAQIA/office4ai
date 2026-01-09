# Office4AI MVP 手动测试指南

本指南提供详细的手动测试步骤，用于验证 Workspace 与 Word Add-In 的端到端通信。

---

## 测试环境要求

### Python 端 (office4ai)
- Python 3.10+ 推荐 3.11
- uv 包管理器
- 所有依赖已安装 (`poe install-dev`)

### TypeScript 端 (office-editor4ai)
- Node.js 18+
- npm 或 yarn
- Word 桌面版 (Microsoft 365 或 Office 2021+)

---

## 测试准备

### 1. 启动 Workspace 服务器

```bash
# 方式 1: 使用健康检查模式 (推荐用于快速测试)
uv run python manual_tests/test_word_e2e.py --mode health

# 方式 2: 使用交互式启动脚本
uv run python manual_tests/test_workspace_startup.py
```

预期输出:
```
============================================================================
✅ Office Workspace started on http://127.0.0.1:3000
============================================================================
Health check: http://127.0.0.1:3000/health
Bind address: 127.0.0.1 (localhost only)
============================================================================
```

### 2. 验证服务器状态

在浏览器中访问: http://127.0.0.1:3000/health

预期响应:
```json
{
  "status": "ok",
  "service": "office4ai-workspace",
  "connections": 0,
  "documents": 0
}
```

### 3. 启动 Word Add-In

```bash
cd /Users/jqq/WebstormProjects/office-editor4ai/word-editor4ai

# 启动开发服务器 (使用 pnpm)
pnpm start

# 或者如果使用 npm
npm start
```

### 4. 加载 Add-In

1. 打开 Word 桌面版
2. 创建新文档或打开现有文档
3. 按照 Word 开发者指南加载 sideload Add-In
4. Task Pane 应该会自动打开

---

## 测试场景 1: 基本连接测试

**目标**: 验证 Add-In 能够成功连接到 Workspace 服务器

### 步骤

1. **查看 Workspace 控制台**
   - 应该看到连接日志
   - 预期输出:
   ```
   [INFO] Client connected: /word#<socket_id>
   [INFO] Handshake received from client: <client_id>
   [INFO] Document registered: <document_uri>
   ```

2. **查看 Add-In 浏览器控制台**
   - 打开开发者工具 (F12)
   - 应该看到连接成功日志
   - 预期输出:
   ```
   [Socket.IO Info] Connected successfully to office4ai workspace
   [Socket.IO Info] Client ID: <client_id>
   [Socket.IO Info] Connection status: connected
   ```

3. **验证健康检查端点**
   ```bash
   curl http://127.0.0.1:3000/health
   ```
   - `connections` 应该 >= 1
   - `documents` 应该 >= 1

### 判定标准

- ✅ **通过**: Workspace 控制台显示连接日志, Add-In 控制台显示连接成功
- ❌ **失败**: 连接超时或错误

---

## 测试场景 2: 获取选中内容

**目标**: 验证 Workspace → Socket.IO → Word Add-In 的完整请求-响应流程

### 步骤

1. **启动端到端测试**
   ```bash
   uv run python manual_tests/test_word_e2e.py --mode e2e
   ```

2. **按照测试提示操作**
   - 测试脚本会提示在 Word 中选中一些文本
   - 在 Word 文档中输入并选中一段文本 (例如: "Hello World")

3. **观察测试输出**

   预期成功输出:
   ```
   [1/5] 启动 Office Workspace...
   ✅ Workspace 启动成功

   [2/5] 等待 Word Add-In 连接...
   ✅ Add-In 已连接

   [3/5] 获取已连接文档...
   ✅ 找到 1 个已连接文档:
      1. <document_uri>

   [4/5] 调用 word:get:selectedContent...
   发送动作: word:get:selectedContent

   [5/5] 验证结果...
   ✅ 动作执行成功
   📝 选中文本内容:
      'Hello World'
   ✅ 成功获取选中文本 (长度: 11)
   ```

### 判定标准

- ✅ **通过**: 测试成功获取选中文本
- ❌ **失败**: 测试超时或返回错误

---

## 测试场景 3: 手动调用 (可选)

**目标**: 手动测试 Socket.IO 事件流

### 步骤

1. **在 Word 文档中选中一些文本**

2. **在 Add-In 浏览器控制台中执行**
   ```javascript
   // 获取 socket client 实例
   const client = window.socketClient;

   // 检查连接状态
   console.log('Connected:', client.isConnected());
   console.log('Document URI:', client.getDocumentUri());

   // 监听响应事件
   client.socket.on('word:get:selectedContent:response', (data) => {
     console.log('Response received:', data);
   });

   // 发送请求
   client.socket.emit('word:get:selectedContent', {
     requestId: 'test-' + Date.now(),
     documentUri: client.getDocumentUri(),
   });
   ```

3. **验证响应**
   - 控制台应该打印响应数据
   - 响应应包含 `success: true` 和选中文本内容

### 判定标准

- ✅ **通过**: 控制台显示响应数据且包含正确内容
- ❌ **失败**: 无响应或响应错误

---

## 常见问题排查

### 问题 1: Add-In 无法连接到 Workspace

**症状**:
- Add-In 控制台显示连接错误
- Workspace 控制台无连接日志

**排查步骤**:
1. 确认 Workspace 服务器正在运行: `curl http://127.0.0.1:3000/health`
2. 检查防火墙设置
3. 检查 Add-In 配置中的服务器 URL (应为 `http://127.0.0.1:3000`)
4. 查看 Add-In 浏览器控制台的网络请求

### 问题 2: 请求-响应超时

**症状**:
- 测试显示 "Timeout waiting for response"
- Workspace 发送请求但未收到响应

**排查步骤**:
1. 检查 Word Add-In 是否正确注册了事件处理器
2. 查看 Add-In 控制台是否有 JavaScript 错误
3. 检查 Workspace 控制台是否显示请求已发送
4. 确认 `requestId` 正确传递和返回

### 问题 3: 文档 URI 未找到

**症状**:
- 测试显示 "No socket found for document"

**排查步骤**:
1. 确认 Add-In 已连接并完成握手
2. 检查 Add-In 控制台的文档 URI 日志
3. 查看 Workspace 控制台的连接管理器日志
4. 尝试重新加载 Add-In

### 问题 4: Word API 调用失败

**症状**:
- Add-In 控制台显示 Word API 错误
- 响应显示 `success: false`

**排查步骤**:
1. 确认 Word 文档已打开且处于活动状态
2. 检查 Add-In 是否有足够的权限
3. 查看 `manifest.xml` 中的权限设置
4. 尝试重新加载 Word 文档

---

## 成功标准

### 最小可行产品 (MVP) 通过条件

- ✅ Workspace 服务器成功启动 (http://127.0.0.1:3000)
- ✅ Word Add-In 成功连接到 Workspace
- ✅ 测试脚本能够获取已连接文档列表
- ✅ 调用 `word:get:selectedContent` 成功返回选中文本
- ✅ 请求-响应机制正常工作 (超时 < 10s)

### 扩展测试 (可选)

- ✅ 测试 `word:insert:text` 插入文本
- ✅ 测试 `word:replace:selection` 替换选中内容
- ✅ 测试多个文档同时连接
- ✅ 测试连接断开后的重连机制

---

## 测试报告模板

```markdown
## Office4AI MVP 测试报告

**测试日期**: YYYY-MM-DD
**测试人员**: [Name]
**环境**:
- office4ai commit: [commit hash]
- office-editor4ai commit: [commit hash]
- Python version: [version]
- Node.js version: [version]
- Word version: [version]

### 测试结果

| 场景 | 状态 | 备注 |
|------|------|------|
| 场景 1: 基本连接测试 | ✅/❌ | [备注] |
| 场景 2: 获取选中内容 | ✅/❌ | [备注] |
| 场景 3: 手动调用 | ✅/❌/N/A | [备注] |

### 问题描述

[如有问题，详细描述]

### 截图/日志

[附上相关截图或日志]

### 结论

- ✅ MVP 通过
- ❌ MVP 未通过
```

---

**最后更新**: 2026-01-05
**维护者**: JQQ <jqq1716@gmail.com>
