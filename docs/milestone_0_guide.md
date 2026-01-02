<!--
文件名: milestone_0_guide.md
作者: JQQ
创建日期: 2025/12/18
最后修改日期: 2025/12/18
版权: 2023 JQQ. All rights reserved.
依赖: None
描述: Office4AI MCP Server Milestone 0 使用指南
-->

# Office4AI MCP Server - Milestone 0 使用指南

## 概述 | Overview

Milestone 0 实现了 Office4AI MCP Server 的基础功能，包括：
- 可启动的 MCP Server（支持 stdio/sse/streamable-http 三种传输模式）
- 基础的 MCP 协议支持（list_tools/list_resources）
- 统一的日志和配置管理

## 快速开始 | Quick Start

### 1. 安装依赖 | Install Dependencies

```bash
uv sync
```

### 2. 启动 MCP Server | Start MCP Server

#### 方式一：使用项目脚本 | Method 1: Using project script
```bash
uv run python scripts/start_mcp_server.py
```

#### 方式二：使用 pyproject.toml 中的命令 | Method 2: Using command in pyproject.toml
```bash
uv run office4ai-mcp
```

### 3. 测试 MCP Server | Test MCP Server

运行完整的测试套件：
```bash
uv run python scripts/test_mcp_client.py
```

## 传输模式 | Transport Modes

### 1. STDIO 模式（默认）| STDIO Mode (Default)
适用于命令行工具和 IDE 集成：
```bash
uv run office4ai-mcp --transport stdio
```

### 2. SSE 模式 | SSE Mode
适用于 Web 应用：
```bash
uv run python scripts/test_sse_transport.py
```
或通过环境变量：
```bash
TRANSPORT=sse uv run office4ai-mcp
```

### 3. Streamable-HTTP 模式 | Streamable-HTTP Mode
适用于 HTTP API 调用：
```bash
uv run python scripts/test_http_transport.py
```
或通过环境变量：
```bash
TRANSPORT=streamable-http uv run office4ai-mcp
```

## 配置选项 | Configuration Options

可以通过环境变量或命令行参数配置：

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `TRANSPORT` | `stdio` | 传输模式：stdio/sse/streamable-http |
| `HOST` | `127.0.0.1` | 服务器监听地址（SSE/HTTP 模式） |
| `PORT` | `8000` | 服务器端口（SSE/HTTP 模式） |

示例：
```bash
# 使用命令行参数
uv run office4ai-mcp --transport sse --host 0.0.0.0 --port 9000

# 使用环境变量
TRANSPORT=sse HOST=0.0.0.0 PORT=9000 uv run office4ai-mcp
```

## 验收标准 | Acceptance Criteria

✅ Milestone 0 已达成以下验收标准：
- [x] MCP Server 可启动（支持 stdio/sse/http 三种传输模式）
- [x] `list_tools` 返回空列表但不报错
- [x] `list_resources` 返回空列表但不报错
- [x] 统一日志、配置读取（MCPServerConfig）
- [x] MCP 客户端能连接并完成 handshake

## 下一步 | Next Steps

Milestone 1 将实现：
- UNO/LibreOffice 进程管理
- 基础文档操作工具（DocOpen/DocCreate/DocSave/DocClose）
- 文档资源（office://document/active）

## 故障排除 | Troubleshooting

### 1. 导入错误 | Import Errors
确保已安装所有依赖：
```bash
uv sync
```

### 2. UNO 相关错误 | UNO Related Errors
Milestone 0 不涉及 UNO 操作，如果遇到 UNO 相关错误，请检查代码是否误用了 Milestone 1+ 的功能。

### 3. 端口占用 | Port Already in Use
更换端口或关闭占用端口的进程：
```bash
PORT=8001 uv run office4ai-mcp
```

## 项目结构 | Project Structure

```
office4ai/
├── office4ai/office/mcp/
│   └── server.py          # MCP Server 主入口
├── office4ai/a2c_smcp/
│   ├── server.py          # BaseMCPServer 基类
│   └── config.py          # 配置管理
├── scripts/
│   ├── start_mcp_server.py    # 启动脚本
│   ├── test_mcp_client.py     # 测试脚本
│   ├── test_sse_transport.py  # SSE 模式测试
│   └── test_http_transport.py # HTTP 模式测试
└── docs/
    └── milestone_0_guide.md   # 本文档
```
