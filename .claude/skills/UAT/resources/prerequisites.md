# UAT 前置条件

## 环境要求

### 1. Office4AI MCP Server

MCP Inspector 以 stdio 模式自动管理 MCP Server 进程，无需手动启动。

### 2. MCP Inspector

```bash
npx @anthropic-ai/mcp-inspector
```

- 默认地址：`http://localhost:6274`
- 确保已连接到 Office4AI MCP Server（左侧显示连接状态）

**注意**：本 UAT 仅做注册验证，不需要 Office Add-In 连接或测试文档。

## MCP Inspector 操作指引

### 查看 Tool 列表

1. 点击左侧 **Tools** 标签页
2. 工具列表中显示所有已注册工具
3. 点击某个工具可查看其参数 schema（inputSchema）

### 查看 Resource 列表

1. 点击左侧 **Resources** 标签页
2. 资源列表中显示所有已注册资源及其 URI、name、mimeType
