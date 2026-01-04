# Office4AI Workspace & Socket.IO 实现计划

> **版本**: 1.1.0
> **创建日期**: 2026-01-03
> **最后更新**: 2026-01-04
> **状态**: MCP Tools 设计完成

---

## 📋 目录

1. [项目背景](#1-项目背景)
2. [架构设计](#2-架构设计)
3. [核心概念](#3-核心概念)
4. [实现阶段](#4-实现阶段)
5. [技术选型](#5-技术选型)
6. [目录结构](#6-目录结构)
7. [MCP 工具设计](#7-mcp-工具设计)
8. [测试策略](#8-测试策略)
9. [风险管理](#9-风险管理)

---

## 1. 项目背景

### 1.1 当前状态

**office4ai** (Python):
- ✅ 已实现 MCP Server 基础设施 (Milestone 0)
- ✅ 支持 stdio/sse/http 三种传输模式
- ✅ 实现了 Gymnasium 环境接口
- ❌ 缺少对 Office Add-In 的管理能力
- ❌ 缺少 Socket.IO 服务器

**office-editor4ai** (TypeScript):
- ✅ 完整的 Office Add-In 实现 (Excel/Word/PPT)
- ✅ 丰富的 Office.js 工具集
- ✅ 完善的 Socket.IO API 标准规范
- ❌ 缺少 Socket.IO 客户端实现
- ❌ 缺少统一的控制逻辑层

### 1.2 核心设计原则

⚠️ **Workspace 作为工作会话（Workspace-as-Session）**

**关键理解**：
1. **Workspace = 工作会话**：类似于 VSCode 的 Workspace，一个 Workspace 可以管理多个文档
2. **Document = 独立连接**：每个打开的文档仍然是一个独立的 Socket.IO 连接
3. **多对多关系**：一个 Document 可以被多个 Workspace 引用，但 Socket.IO 连接只有一个

**类比 VSCode**：
```
VSCode Workspace "MyProject":
├── file1.py
├── file2.py
└── file3.py

Office4AI Workspace "Report Project":
├── report.docx (Word)
├── data.xlsx (Excel)
└── presentation.pptx (PPT)
```

**⚠️ 关键场景：一个文档，多个工作区**：
```
用户场景：
- 打开了 report.docx（只有 1 个 Socket.IO 连接）
- 创建了 Workspace A "报告项目"，添加 report.docx
- 创建了 Workspace B "文档整理"，也添加 report.docx

架构设计：
- 1 个 Document（report.docx, client_id: word-abc123）
- 2 个 Workspace（A 和 B）都引用同一个 Document
- Socket.IO 连接只有 1 个，但被 2 个 Workspace 共享
- 当 report.docx 关闭时，需要从 2 个 Workspace 中都移除该引用
```

**设计影响**：
1. **多对多关系**：
   - `Workspace.documents: list[Document]` - 一个 Workspace 包含多个文档
   - `Document.workspace_ids: list[str]` - 一个文档可以被多个 Workspace 引用
2. **Socket.IO 连接独立于 Workspace**：每个文档只有 1 个 Socket.IO 连接，但可被多个 Workspace 共享
3. **生命周期管理**：文档关闭时，需要从所有引用它的 Workspace 中移除
4. **MCP 工具可以指定操作层级**：
   - `workspace_id`：操作整个 Workspace（如列出所有文档）
   - `document_uri`：操作特定文档（可能影响多个 Workspace）

### 1.3 整体架构图

```
┌─────────────────────────────────────────────────────────────┐
│                     AI Agent / User                         │
└──────────────────────────┬──────────────────────────────────┘
                           │ MCP Protocol (stdio/sse/http)
                           ↓
┌────────────────────────────────────────────────────────────┐
│              Office4AI MCP Server (Python)                 │
│                                                            │
│  ┌────────────────────────────────────────────────────┐    │
│  │  MCP Tools Layer                                   │    │
│  │  - workspace:create                                │    │
│  │  - workspace:list                                  │    │
│  │  - word:get:selectedContent                        │    │
│  │  - ppt:insert:text                                 │    │
│  │  - excel:insert:table                              │    │
│  └──────────────────┬─────────────────────────────────┘    │
│                     │                                      │
│  ┌──────────────────▼─────────────────────────────────┐    │
│  │  Workspace Manager                                 │    │
│  │  - Track active workspaces                         │    │
│  │  - Manage document lifecycle                       │    │
│  │  - Route commands to correct socket                │    │
│  └──────────────────┬─────────────────────────────────┘    │
│                     │                                      │
│  ┌──────────────────▼─────────────────────────────────┐    │
│  │  Socket.IO Server (python-socketio)                │    │
│  │  - /word namespace                                 │    │
│  │  - /ppt namespace                                  │    │
│  │  - /excel namespace                                │    │
│  │  - Control Service (business logic)                │    │
│  └───────────────────┬────────────────────────────────┘    │
└──────────────────────┼─────────────────────────────────────┘
                       │ Socket.IO (WebSocket)
                       ↓
┌──────────────────────────────────────────────────────────┐
│              Office Editor4AI Add-Ins                    │
│                                                          │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐    │
│  │ Word Add-In  │  │ PPT Add-In   │  │ Excel Add-In │    │
│  │              │  │              │  │              │    │
│  │ Socket.IO    │  │ Socket.IO    │  │ Socket.IO    │    │
│  │ Client       │  │ Client       │  │ Client       │    │
│  │              │  │              │  │              │    │
│  │ Office.js    │  │ Office.js    │  │ Office.js    │    │
│  │ Tools        │  │ Tools        │  │ Tools        │    │
│  └──────────────┘  └──────────────┘  └──────────────┘    │
└──────────────────────────────────────────────────────────┘
                       │
                       ↓
              ┌──────────────────┐
              │  Office Apps     │
              │  (Word/PPT/Excel)│
              └──────────────────┘
```

---

## 2. 架构设计

### 2.1 三层架构模式

参考 `examples/ide4ai` 项目，采用清晰的三层架构：

```
┌─────────────────────────────────────────────────────┐
│   Layer 3: Specific Implementation                  │
│   (office/word/, office/excel/, office/pptx/)       │
│   - WordOfficeEnv, ExcelOfficeEnv, PPTXOfficeEnv    │
│   - WordDocumentEnv, ExcelDocumentEnv, PPTXEnv     │
│   - Office 特定 MCP Tools                           │
└─────────────────────────────────────────────────────┘
                        ▲ 继承/扩展
                        │
┌─────────────────────────────────────────────────────┐
│   Layer 2: Generic Infrastructure                   │
│   (a2c_smcp/, environment/, socketio/)              │
│   - BaseMCPServer, BaseTool, BaseResource           │
│   - BaseDocumentEnv, BaseTerminalEnv                │
│   - Socket.IO Client (A2C) + Server (Add-In)        │
└─────────────────────────────────────────────────────┘
                        △ 依赖/使用
                        │
┌─────────────────────────────────────────────────────┐
│   Layer 1: Core Foundation                          │
│   (office4ai/ 根目录)                               │
│   - OfficeEnv (base.py)                             │
│   - OfficeAction, OfficeObs (schema.py)             │
│   - OfficeExecutionError (exceptions.py)            │
│   - OfficeSingleton (singleton.py)                  │
│   - dtos/ (数据传输对象)                             │
└─────────────────────────────────────────────────────┘
```

**架构说明**：
- **Layer 1（核心层）**：定义基础接口、数据模型、异常处理，100% 通用
- **Layer 2（基础设施层）**：实现 MCP 协议、环境抽象、Socket.IO 双重角色，80% 通用
- **Layer 3（特定实现层）**：Office 特定功能（Word/Excel/PPT），20% 通用

### 2.2 职责分离

| 层级 | 职责 | 通用性 | 示例 |
|------|------|--------|------|
| **核心层** | 定义基础接口、数据模型、异常处理 | 100% 通用 | `OfficeEnv`, `OfficeAction`, `OfficeObs` |
| **基础设施层** | 实现 MCP 协议、环境抽象、Socket.IO 通信 | 80% 通用 | `BaseMCPServer`, `BaseDocumentEnv`, `A2CClient` |
| **特定实现层** | Office 特定功能 | 20% 通用 | `WordOfficeEnv`, `DocxEditTool`, `ExcelGetRangeTool` |

### 2.3 Socket.IO 双重角色（关键架构）

**⚠️ 核心理解**：office4ai 在 Socket.IO 通信中扮演**两个角色**：

```
┌──────────────────────┐
│   A2C Server         │
│   (Computer)         │
└──────────┬───────────┘
           │ Socket.IO
           │ (office4ai 作为 Client)
           ↓
┌──────────────────────────────────────────────────────┐
│              office4ai                                │
│                                                        │
│  ┌────────────────────────────────────────────────┐  │
│  │  Role 1: A2C Computer Client                   │  │
│  │  - 连接到 A2C Server                          │  │
│  │  - 接收计算机控制指令                          │  │
│  │  - 上报环境状态                               │  │
│  └────────────────────────────────────────────────┘  │
│                                                        │
│  ┌────────────────────────────────────────────────┐  │
│  │  Role 2: Socket.IO Server (管理 Add-In)        │  │
│  │  - 启动 Socket.IO 服务器                       │  │
│  │  - 监听 /word, /ppt, /excel namespace          │  │
│  │  - 管理多个 Add-In 客户端连接                  │  │
│  │  - 向 Add-In 发送操作指令                      │  │
│  │  - 接收 Add-In 事件上报                        │  │
│  └────────────────────────────────────────────────┘  │
│                                                        │
└──────────────────────┬───────────────────────────────┘
                       │ Socket.IO
                       │ (office4ai 作为 Server)
                       ↓
┌──────────────────────────────────────────────────────┐
│         Office Add-In Clients (Word/Excel/PPT)       │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐│
│  │ Word Add-In  │  │ Excel Add-In │  │ PPT Add-In   ││
│  │              │  │              │  │              ││
│  │ Socket.IO    │  │ Socket.IO    │  │ Socket.IO    ││
│  │ Client       │  │ Client       │  │ Client       ││
│  └──────────────┘  └──────────────┘  └──────────────┘│
└──────────────────────────────────────────────────────┘
         │
         ↓
┌──────────────────┐
│  Office Apps     │
│  (Word/Excel/PPT) │
└──────────────────┘
```

**关键点**：
1. **对外**：office4ai 是 **A2C Computer Client**，连接到 A2C Server
2. **对内**：office4ai 是 **Socket.IO Server**，管理多个 Office Add-In
3. **双向通信**：
   - A2C Server ←→ office4ai (Client-Server 模式)
   - office4ai ←→ Office Add-In (Server-Client 模式)

---

## 7. MCP 工具设计

### 7.1 设计决策总结

基于对 `examples/ide4ai` 项目和 `office-editor4ai` Socket.IO API 规范的深入分析，确定以下设计决策：

| 决策项 | 选择 | 理由 |
|--------|------|------|
| **工具组织方式** | 按 Office 类型分组 | 与 Socket.IO API 对齐，工具名称清晰明确 |
| **工具粒度** | 中等（27 个工具） | 覆盖 80% 使用场景，精简高效 |
| **文档定位方式** | 每次传入 document_uri | 无状态设计，简单直接，易于理解 |
| **批量操作** | 不支持 | 简化实现，AI Agent 可通过多次调用实现批量效果 |
| **错误处理** | 严格模式 | 参考 ide4ai，多重验证机制（Schema + 业务逻辑 + 状态检查） |
| **Schema 驱动** | Pydantic BaseModel | 自动生成 JSON Schema，类型安全 |

### 7.2 工具列表（27 个）

采用 `<platform>_<action>[:<resource>]` 命名规范，按 Office 类型分组。

#### 7.2.1 Word Tools (13 个)

**内容获取类（4 个）**

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `word_get_selected_content` | 获取选中内容 | document_uri, options (includeText/includeImages/includeTables...) | text, elements[], metadata |
| `word_get_visible_content` | 获取可见内容 | document_uri, options | text, elements[], stats |
| `word_get_document_structure` | 获取文档结构 | document_uri | sections[], paragraphs[], headings[], tables[] |
| `word_get_document_stats` | 获取统计信息 | document_uri | wordCount, paragraphCount, tableCount, imageCount |

**文本操作类（4 个）**

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `word_insert_text` | 插入文本 | document_uri, text, location (Cursor/Start/End), format? | inserted, position |
| `word_replace_selection` | 替换选中内容（参考 ide4ai Edit） | document_uri, old_string, new_string, replace_all? | replacements_made, undo_info |
| `word_replace_text` | 查找替换文本 | document_uri, search_text, replace_text, options? | replacements_made |
| `word_append_text` | 追加文本 | document_uri, text, location (Start/End) | appended, position |

**多模态操作类（3 个）**

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `word_insert_image` | 插入图片 | document_uri, image_base64, options (width/height/alt_text) | inserted, image_id |
| `word_insert_table` | 插入表格 | document_uri, rows, cols, data?, options? | inserted, table_id |
| `word_insert_equation` | 插入公式 | document_uri, latex, options? | inserted, equation_id |

**高级功能类（2 个）**

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `word_insert_toc` | 插入目录 | document_uri, options? | inserted |
| `word_export_content` | 导出内容 | document_uri, format (text/markdown/html) | content |

#### 7.2.2 PowerPoint Tools (10 个)

**内容获取类（3 个）**

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `ppt_get_current_slide_elements` | 获取当前页元素 | document_uri | elements[], slide_index |
| `ppt_get_slide_elements` | 获取指定页元素 | document_uri, slide_index | elements[], slide_info |
| `ppt_get_slide_screenshot` | 获取幻灯片截图 | document_uri, slide_index? | image_base64, dimensions |

**内容插入类（4 个）**

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `ppt_insert_text` | 插入文本 | document_uri, text, position (left/top), slide_index? | inserted, text_box_id |
| `ppt_insert_image` | 插入图片 | document_uri, image_base64, position, slide_index? | inserted, image_id |
| `ppt_insert_table` | 插入表格 | document_uri, rows, cols, position, slide_index? | inserted, table_id |
| `ppt_insert_shape` | 插入形状 | document_uri, shape_type, position, options?, slide_index? | inserted, shape_id |

**幻灯片管理类（2 个）**

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `ppt_delete_slide` | 删除幻灯片 | document_uri, slide_index | deleted |
| `ppt_move_slide` | 移动幻灯片 | document_uri, from_index, to_index | moved |

**更新操作类（1 个）**

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `ppt_update_text_box` | 更新文本框 | document_uri, text_box_id, new_text, slide_index? | updated |

#### 7.2.3 Excel Tools (4 个)

| 工具名称 | 功能描述 | 输入参数 | 输出 |
|---------|---------|---------|------|
| `excel_get_selected_range` | 获取选中范围 | document_uri | range, values, formulas |
| `excel_set_cell_value` | 设置单元格值 | document_uri, cell_address, value, format? | updated |
| `excel_insert_table` | 插入表格 | document_uri, range, data? | inserted, table_id |
| `excel_get_used_range` | 获取已使用范围 | document_uri | range, row_count, column_count |

### 7.3 工具对比分析

| 类别 | Socket.IO API | MCP Tools | 精简说明 |
|------|---------------|-----------|----------|
| **Word** | 26 个 | 13 个 | -13 (合并相似功能) |
| **PowerPoint** | 21 个 | 10 个 | -11 (合并更新操作) |
| **Excel** | 待补充 | 4 个 | 核心功能 |
| **总计** | 47+ | 27 个 | 覆盖 80% 使用场景 |

**精简策略**：
- 合并 `insert_text_to_cursor` 和 `insert_text_to_end` → `insert_text` + `location` 参数
- 合并 `update_text_box` 和 `update_text_boxes_batch` → 仅保留单个更新（不需要批量）
- 移除 `get_header_footer`、`get_comments` 等低频工具
- Excel 仅保留核心功能，后续扩展

### 7.4 Schema 设计

#### 7.4.1 基础 Schema

```python
# office4ai/a2c_smcp/schemas/common.py
from pydantic import BaseModel, Field
from typing import Any

class OfficeToolOutput(BaseModel):
    """统一的工具输出结构"""
    success: bool = Field(..., description="操作是否成功")
    message: str = Field(default="", description="面向人类的操作结果说明")
    data: dict[str, Any] = Field(default_factory=dict, description="结构化结果数据")
    error_code: str | None = Field(default=None, description="错误码")
    error_message: str | None = Field(default=None, description="错误详情")
    metadata: dict[str, Any] = Field(
        default_factory=dict,
        description="元数据（执行耗时、文档版本等）"
    )

class DocumentTarget(BaseModel):
    """文档定位参数（所有工具的公共参数）"""
    document_uri: str = Field(..., description="文档 URI（绝对路径或 file:// URI）")
```

#### 7.4.2 Word Tools Schema 示例

```python
# office4ai/a2c_smcp/schemas/word.py

class WordGetSelectedContentInput(BaseModel):
    """word_get_selected_content 输入"""
    document_uri: str = Field(..., description="Word 文档 URI")
    include_text: bool = Field(default=True, description="是否包含文本内容")
    include_images: bool = Field(default=True, description="是否包含图片信息")
    include_tables: bool = Field(default=True, description="是否包含表格信息")
    max_text_length: int | None = Field(default=None, description="最大文本长度")

class WordInsertTextInput(BaseModel):
    """word_insert_text 输入"""
    document_uri: str = Field(..., description="Word 文档 URI")
    text: str = Field(..., description="要插入的文本")
    location: Literal["Cursor", "Start", "End"] = Field(
        default="Cursor",
        description="插入位置"
    )
    format: dict | None = Field(
        default=None,
        description="文本格式 {bold, italic, fontSize, fontName, color}"
    )

class WordReplaceSelectionInput(BaseModel):
    """word_replace_selection 输入（参考 ide4ai Edit）"""
    document_uri: str = Field(..., description="Word 文档 URI")
    old_string: str = Field(..., description="要替换的文本")
    new_string: str = Field(..., description="替换后的文本")
    replace_all: bool = Field(
        default=False,
        description="是否替换所有匹配项（如果 old_string 不唯一）"
    )
```

#### 7.4.3 PowerPoint Tools Schema 示例

```python
# office4ai/a2c_smcp/schemas/ppt.py

class PptInsertTextInput(BaseModel):
    """ppt_insert_text 输入"""
    document_uri: str = Field(..., description="PowerPoint 文档 URI")
    text: str = Field(..., description="要插入的文本")
    left: float = Field(..., description="文本框左边距（英寸）")
    top: float = Field(..., description="文本框上边距（英寸）")
    width: float | None = Field(default=None, description="文本框宽度")
    height: float | None = Field(default=None, description="文本框高度")
    slide_index: int | None = Field(default=None, description="目标幻灯片索引（默认当前页）")
```

### 7.5 工具实现架构

#### 7.5.1 目录结构（对齐 ide4ai 三层架构）

```
office4ai/                              # 主包
│
├── [Layer 1: Core Foundation]
├── base.py                             # ⭐ OfficeEnv 基类 (gym.Env)
├── schema.py                           # ⭐ OfficeAction, OfficeObs
├── exceptions.py                       # ⭐ OfficeExecutionError 等
├── utils.py                            # ⭐ 工具函数
├── singleton.py                        # ⭐ OfficeSingleton (单例管理)
│
├── dtos/                               # 数据传输对象
│   ├── __init__.py
│   ├── base_protocol.py                # Socket.IO 协议定义
│   ├── document.py                     # 文档 DTO
│   └── commands.py                     # 命令 DTO
│
├── [Layer 2: Generic Infrastructure]
├── a2c_smcp/                           # MCP 基础设施
│   ├── server.py                       # BaseMCPServer
│   ├── config.py                       # MCPServerConfig
│   │
│   ├── schemas/                        # ⭐ 工具 Schema
│   │   ├── __init__.py
│   │   ├── common.py                   # OfficeToolOutput, DocumentTarget
│   │   ├── word.py                     # Word 工具 Schema
│   │   ├── ppt.py                      # PowerPoint 工具 Schema
│   │   └── excel.py                    # Excel 工具 Schema
│   │
│   └── tools/                          # 工具（按应用分类）
│       ├── __init__.py
│       ├── base.py                     # BaseTool 基类
│       │
│       ├── word/                       # ⭐ Word 工具目录
│       │   ├── __init__.py
│       │   ├── get_selected_content.py
│       │   ├── insert_text.py
│       │   ├── replace_selection.py
│       │   └── ...
│       │
│       ├── ppt/                        # ⭐ PPT 工具目录
│       │   ├── __init__.py
│       │   ├── insert_text.py
│       │   └── ...
│       │
│       └── excel/                      # ⭐ Excel 工具目录
│           ├── __init__.py
│           └── ...
│
├── environment/                        # 环境组件
│   ├── terminal/                       # ✅ 复用 ide4ai
│   │   ├── base.py
│   │   ├── local_terminal_env.py
│   │   └── pexpect_terminal_env.py
│   │
│   └── workspace/                      # ⭐ 文档工作空间（对应 ide4ai workspace）
│       ├── __init__.py
│       ├── base.py                     # BaseWorkspaceEnv
│       ├── schema.py                   # DocumentRange, Position 等
│       └── model.py                    # DocumentModel
│
├── socketio/                           # ⭐ Socket.IO 双重角色
│   ├── __init__.py
│   │
│   ├── client/                         # ⭐ Role 1: A2C Client
│   │   ├── __init__.py
│   │   ├── a2c_client.py               # 连接到 A2C Server
│   │   ├── handlers.py                 # A2C 协议处理器
│   │   └── config.py                   # A2C 连接配置
│   │
│   └── server/                         # ⭐ Role 2: Socket.IO Server (管理 Add-In)
│       ├── __init__.py
│       ├── server.py                   # Socket.IO 服务器主入口
│       ├── namespaces/                 # 命名空间实现
│       │   ├── __init__.py
│       │   ├── word.py                 # Word namespace
│       │   ├── ppt.py                  # PowerPoint namespace
│       │   └── excel.py                # Excel namespace
│       ├── middleware/                  # 中间件
│       │   ├── auth.py                 # 认证中间件
│       │   └── logging.py              # 日志中间件
│       └── services/                    # 控制服务
│           ├── connection_manager.py    # 连接管理
│           ├── event_dispatcher.py      # 事件分发
│           └── cache.py                 # 缓存服务
│
└── [Layer 3: Specific Implementation]
└── office/                             # Office 特定实现
    ├── __init__.py
    │
    ├── base.py                         # ⭐ Office 基类
    │   # class OfficeEnv(gym.Env)
    │
    ├── session.py                      # ⭐ AddInSession 接口
    │   # class AddInSession(ABC)
    │   #   - async def call_word_addin(...)
    │   #   - async def call_ppt_addin(...)
    │   #   - async def call_excel_addin(...)
    │
    ├── word/                           # ⭐ Word 特定实现
    │   ├── __init__.py
    │   ├── env.py                       # WordOfficeEnv (OfficeEnv 子类)
    │   ├── workspace_env.py            # WordWorkspaceEnv (BaseWorkspaceEnv 子类)
    │   └── a2c_smcp/                   # Word MCP Server（可选）
    │       ├── __init__.py
    │       └── server.py               # WordMCPServer (可选)
    │
    ├── excel/                          # ⭐ Excel 特定实现
    │   ├── __init__.py
    │   ├── env.py                       # ExcelOfficeEnv
    │   ├── workspace_env.py            # ExcelWorkspaceEnv
    │   └── a2c_smcp/
    │       ├── __init__.py
    │       └── server.py               # ExcelMCPServer (可选)
    │
    ├── pptx/                           # ⭐ PowerPoint 特定实现
    │   ├── __init__.py
    │   ├── env.py                       # PPTXOfficeEnv
    │   ├── workspace_env.py            # PPTXWorkspaceEnv
    │   └── a2c_smcp/
    │       ├── __init__.py
    │       └── server.py               # PPTXMCPServer (可选)
    │
    ├── mcp/                            # Office MCP Server
    │   ├── __init__.py
    │   └── server.py                   # OfficeMCPServer
    │   #   def _register_tools(self)    # ⭐ 需实现
    │   #   def _register_resources(self) # ⭐ 需实现
    │
    └── legacy/                         # 遗留代码（如果需要）
        ├── __init__.py
        ├── uno_bridge/                 # UNO Bridge（LibreOffice 集成）
        └── handlers/                   # 旧版处理器
            ├── docx_handler.py
            ├── pptx_handler.py
            └── xlsx_handler.py
```

**架构特点**：
- **Layer 1**：核心基类和 Schema（OfficeEnv, OfficeAction/Obs, OfficeExecutionError）
- **Layer 2**：通用基础设施（MCP 协议、环境抽象（terminal + workspace）、Socket.IO 双重角色）
- **Layer 3**：Office 特定实现（word/excel/pptx 目录）
- **命名统一**：全面使用 "Office" 术语，不再使用 "IDE"
- **环境对齐**：使用 `workspace/` 对齐 ide4ai 的 workspace 概念（虽然 office4ai 操作的是文档）
- **Socket.IO 双重角色**：
  - `socketio/client/` - A2C Client 角色
  - `socketio/server/` - Socket.IO Server 角色（管理 Add-In）

#### 7.5.2 工具基类设计

```python
# office4ai/a2c_smcp/tools/base.py
class BaseOfficeTool(BaseTool):
    """Office 工具基类"""

    def __init__(self, session: "DocumentSession"):
        self.session = session

    def validate_input(self, arguments: dict, model: type[T]) -> T:
        """参数验证（继承自 BaseTool）"""
        return model.model_validate(arguments)

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """执行工具（子类实现）"""
        raise NotImplementedError
```

#### 7.5.3 DocumentSession 接口设计

```python
# office4ai/office/session.py
from abc import ABC, abstractmethod

class DocumentSession(ABC):
    """文档会话接口（负责 Socket.IO 通信）"""

    @abstractmethod
    async def call_word_tool(
        self,
        event: str,
        data: dict[str, Any]
    ) -> dict[str, Any]:
        """调用 Word 工具"""
        pass

    @abstractmethod
    async def call_ppt_tool(
        self,
        event: str,
        data: dict[str, Any]
    ) -> dict[str, Any]:
        """调用 PowerPoint 工具"""
        pass

    @abstractmethod
    async def call_excel_tool(
        self,
        event: str,
        data: dict[str, Any]
    ) -> dict[str, Any]:
        """调用 Excel 工具"""
        pass
```

#### 7.5.4 工具实现示例

```python
# office4ai/a2c_smcp/tools/word/insert_text.py
class WordInsertTextTool(BaseOfficeTool):
    @property
    def name(self) -> str:
        return "word_insert_text"

    @property
    def description(self) -> str:
        return """在 Word 文档中插入文本。

        使用说明：
        - 在文档中指定位置插入文本
        - 可以指定文本格式（粗体、斜体、字号等）
        - 插入位置：Cursor（光标处）、Start（文档开头）、End（文档结尾）
        """

    @property
    def input_schema(self) -> dict[str, Any]:
        from office4ai.a2c_smcp.schemas.word import WordInsertTextInput
        return WordInsertTextInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        from office4ai.a2c_smcp.schemas.word import (
            WordInsertTextInput,
            WordInsertTextOutput
        )

        # 1. 参数验证
        try:
            input_data = self.validate_input(arguments, WordInsertTextInput)
        except ValidationError as e:
            return OfficeToolOutput(
                success=False,
                error_code="VALIDATION_ERROR",
                error_message=str(e)
            ).model_dump()

        # 2. 调用 DocumentSession
        try:
            response = await self.session.call_word_tool(
                event="word:insert:text",
                data={
                    "requestId": str(uuid.uuid4()),
                    "text": input_data.text,
                    "location": input_data.location,
                    "format": input_data.format
                }
            )

            # 3. 构造输出
            if response.get("success"):
                return WordInsertTextOutput(
                    success=True,
                    message=f"成功插入文本到 {input_data.location}",
                    data=response.get("data", {}),
                    metadata={"duration": response.get("duration")}
                ).model_dump()
            else:
                return WordInsertTextOutput(
                    success=False,
                    error_code=response.get("error", {}).get("code"),
                    error_message=response.get("error", {}).get("message")
                ).model_dump()

        except Exception as e:
            return OfficeToolOutput(
                success=False,
                error_code="UNKNOWN_ERROR",
                error_message=str(e)
            ).model_dump()
```

#### 7.5.5 工具注册

```python
# office4ai/office/mcp/server.py
class OfficeMCPServer(BaseMCPServer):
    def __init__(
        self,
        config: MCPServerConfig,
        document_session: DocumentSession
    ):
        super().__init__(config, "Office4AI")
        self.document_session = document_session
        self._register_tools()

    def _register_tools(self) -> None:
        from office4ai.a2c_smcp.tools.word import (
            WordGetSelectedContentTool,
            WordInsertTextTool,
            WordReplaceSelectionTool,
            # ... 其他 Word 工具
        )
        from office4ai.a2c_smcp.tools.ppt import (
            PptInsertTextTool,
            PptGetSlideElementsTool,
            # ... 其他 PPT 工具
        )
        from office4ai.a2c_smcp.tools.excel import (
            ExcelGetSelectedRangeTool,
            ExcelSetCellValueTool,
            # ... 其他 Excel 工具
        )

        # Word 工具
        self.tools["word_get_selected_content"] = \
            WordGetSelectedContentTool(self.document_session)
        self.tools["word_insert_text"] = \
            WordInsertTextTool(self.document_session)
        # ... 注册其他工具

        # PPT 工具
        self.tools["ppt_insert_text"] = \
            PptInsertTextTool(self.document_session)
        # ... 注册其他工具

        # Excel 工具
        self.tools["excel_get_selected_range"] = \
            ExcelGetSelectedRangeTool(self.document_session)
        # ... 注册其他工具
```

### 7.6 实现阶段

#### 阶段 1：基础设施（优先级最高）
1. 实现 `OfficeToolOutput`、`DocumentTarget` 基础 Schema
2. 实现 `DocumentSession` 接口和 Socket.IO 客户端
3. 实现 `BaseOfficeTool` 基类

#### 阶段 2：Word Tools（第一批）
1. `word_get_selected_content` - 获取选中内容
2. `word_insert_text` - 插入文本
3. `word_replace_selection` - 替换选中内容（核心参考 ide4ai Edit）
4. `word_insert_image` - 插入图片

#### 阶段 3：PowerPoint Tools（第二批）
1. `ppt_get_current_slide_elements` - 获取当前页元素
2. `ppt_insert_text` - 插入文本
3. `ppt_insert_image` - 插入图片
4. `ppt_delete_slide` - 删除幻灯片

#### 阶段 4：Excel Tools（第三批）
1. `excel_get_selected_range` - 获取选中范围
2. `excel_set_cell_value` - 设置单元格值

#### 阶段 5：完善和测试
1. 补充剩余工具
2. 编写单元测试和集成测试
3. 完善错误处理和日志

### 7.7 关键文件清单

需要创建/修改的文件：

#### 新建文件
1. `office4ai/a2c_smcp/schemas/__init__.py`
2. `office4ai/a2c_smcp/schemas/common.py` - 基础 Schema
3. `office4ai/a2c_smcp/schemas/word.py` - Word Schema
4. `office4ai/a2c_smcp/schemas/ppt.py` - PPT Schema
5. `office4ai/a2c_smcp/schemas/excel.py` - Excel Schema
6. `office4ai/a2c_smcp/tools/word/__init__.py` - Word 工具模块
7. `office4ai/a2c_smcp/tools/word/get_selected_content.py`
8. `office4ai/a2c_smcp/tools/word/insert_text.py`
9. `office4ai/a2c_smcp/tools/word/replace_selection.py`
10. `office4ai/a2c_smcp/tools/ppt/__init__.py` - PPT 工具模块
11. `office4ai/a2c_smcp/tools/ppt/insert_text.py`
12. `office4ai/office/session.py` - DocumentSession 接口
13. `office4ai/office/socket_client.py` - Socket.IO 客户端实现

#### 修改文件
1. `office4ai/office/mcp/server.py` - OfficeMCPServer 工具注册

### 7.8 设计参考

本设计参考了以下项目和规范：

1. **examples/ide4ai** - MCP Tools 封装模式
   - Schema 驱动设计（Pydantic BaseModel）
   - IDE 实例注入（依赖倒置）
   - 统一输出结构（success/message/data/error/metadata）
   - 多重验证机制

2. **office-editor4ai** - Socket.IO API 规范
   - 事件命名规范：`<platform>:<action>[:<resource>]`
   - 47 个可直接映射的工具
   - 统一的请求/响应格式
   - Office.js 工具集完整实现

3. **关键差异与适配**
   - Workspace 无关：工具直接操作 Document，不依赖 Workspace 概念
   - 无状态设计：每次传入 document_uri，而非维护活动文档
   - 精简工具：从 47 个精简到 27 个，覆盖 80% 使用场景
   - 不支持批量操作：简化实现，AI Agent 可多次调用

---

### 7.8 与 ide4ai 架构对齐

#### 7.8.1 命名映射

| ide4ai | office4ai | 说明 |
|--------|-----------|------|
| `IDE` | `OfficeEnv` | 办公环境基类 |
| `IDESingleton` | `OfficeSingleton` | 单例管理 |
| `PythonIDE` | `WordOfficeEnv` | Word 特定环境 |
| `IDEAction` | `OfficeAction` | 动作 Schema |
| `IDEObs` | `OfficeObs` | 观察 Schema |
| `IDEExecutionError` | `OfficeExecutionError` | 执行异常 |
| `BaseWorkspace` | `BaseWorkspaceEnv` | 工作空间环境基类 |
| `LSPWorkspaceEnv` | `WordWorkspaceEnv` | Word 工作空间环境 |

#### 7.8.2 架构对齐度

| 架构层面 | ide4ai | office4ai | 对齐度 |
|---------|--------|-----------|--------|
| **核心层** | `IDE`, `IDEAction/Obs`, `exceptions.py` | `OfficeEnv`, `OfficeAction/Obs`, `exceptions.py` | ✅ 100% |
| **MCP 协议** | `BaseMCPServer`, `BaseTool` | `BaseMCPServer`, `BaseTool` | ✅ 100% |
| **Schema 驱动** | Pydantic BaseModel | Pydantic BaseModel | ✅ 100% |
| **工具组织** | `python_ide/a2c_smcp/tools/` | `office/word/a2c_smcp/tools/` | ✅ 90% |
| **环境抽象** | `workspace/`, `terminal/` | `workspace/`, `terminal/` | ✅ 100% |
| **多重验证** | Schema + 业务逻辑 + 状态检查 | Schema + 业务逻辑 + 状态检查 | ✅ 100% |
| **Socket.IO** | 不涉及 | A2C Client + Add-In Server | ⚠️ N/A (扩展能力) |

#### 7.8.3 工具实现模式对齐

**ide4ai Edit 工具模式**（参考）：
```python
class EditTool(BaseTool):
    async def execute(self, arguments: dict) -> dict:
        # 1. 参数验证（Pydantic）
        input_data = EditInput.model_validate(arguments)

        # 2. 业务逻辑验证
        if not self._is_file_accessible(input_data.file_uri):
            raise ValueError("File not accessible")

        # 3. 状态检查（文件是否被修改）
        if self._is_file_modified(input_data.file_uri):
            raise ValueError("File has been modified")

        # 4. 唯一性检查（old_string 在文件中唯一）
        occurrences = self._count_occurrences(input_data.file_uri, input_data.old_string)
        if occurrences > 1 and not input_data.replace_all:
            raise ValueError(f"String appears {occurrences} times")

        # 5. 执行操作
        result = await self._replace_text(...)
        return EditOutput(success=True, data=result)
```

**office4ai word_replace_selection 工具**（对齐实现）：
```python
class WordReplaceSelectionTool(BaseOfficeTool):
    async def execute(self, arguments: dict) -> dict:
        # 1. 参数验证（Pydantic）
        input_data = WordReplaceSelectionInput.model_validate(arguments)

        # 2. 业务逻辑验证
        if not self._is_document_open(input_data.document_uri):
            raise ValueError("Document not open")

        # 3. 状态检查（文档是否被外部修改）
        if await self._is_document_modified(input_data.document_uri):
            raise ValueError("Document has been modified externally")

        # 4. 唯一性检查（old_string 在选中内容中唯一）
        selection = await self.session.call_word_tool("word:get:selectedContent", ...)
        occurrences = selection["text"].count(input_data.old_string)
        if occurrences > 1 and not input_data.replace_all:
            raise ValueError(f"String appears {occurrences} times in selection")

        # 5. 执行操作
        result = await self.session.call_word_tool("word:replace:text", ...)
        return WordReplaceSelectionOutput(success=True, data=result)
```

#### 7.8.4 扩展能力

office4ai 在对齐 ide4ai 架构的基础上，增加了以下扩展能力：

1. **Socket.IO 双重角色**：
   - **A2C Client**：作为 A2C Computer Client 连接到 A2C Server
   - **Socket.IO Server**：作为服务器管理多个 Office Add-In

2. **多 Office 类型支持**：
   - ide4ai：仅支持 Python IDE
   - office4ai：支持 Word、Excel、PowerPoint 三种 Office 应用

3. **工作空间环境抽象**：
   - ide4ai：基于 LSP Workspace（代码项目）
   - office4ai：基于 Workspace（Office 文档），非 LSP 协议，直接操作 Office.js

---

## 8. 测试策略

### 8.1 单元测试

- Schema 验证测试
- 工具逻辑测试（Mock DocumentSession）
- 错误处理测试

### 8.2 集成测试

- Socket.IO 通信测试
- 端到端工具调用测试
- 与 office-editor4ai 联调测试

### 8.3 测试组织

```
tests/
├── unit_tests/
│   ├── test_schemas.py          # Schema 测试
│   ├── test_tools.py            # 工具单元测试
│   └── test_session.py          # DocumentSession 测试
└── integration_tests/
    ├── test_socketio_client.py  # Socket.IO 客户端测试
    └── test_tools_integration.py # 工具集成测试
```

---

## 9. 风险管理

### 9.1 技术风险

| 风险 | 影响 | 缓解措施 |
|------|------|---------|
| Socket.IO 客户端实现复杂度 | 高 | 复用 python-socketio 库，参考 office-editor4ai 客户端实现 |
| Office.js API 版本兼容性 | 中 | 版本锁定，提供兼容性测试 |
| 文档定位机制设计 | 中 | 采用无状态设计（document_uri），简化实现 |

### 9.2 实施风险

| 风险 | 影响 | 缓解措施 |
|------|------|---------|
| 工具数量庞大，实现周期长 | 中 | 分阶段实施，优先实现核心工具 |
| Schema 设计变更 | 低 | 使用 Pydantic，便于重构和扩展 |
| 测试覆盖不足 | 中 | 编写完整的单元测试和集成测试 |

---

## 附录

### A. 参考文档

- [examples/ide4ai](../../examples/ide4ai) - IDE MCP Server 参考实现
- [office-editor4ai Socket.IO API 规范](https://github.com/your-org/office-editor4ai/blob/main/docs/SOCKET_IO_API_STANDARD.md)
- [MCP Protocol 规范](https://modelcontextprotocol.io)
- [Pydantic 文档](https://docs.pydantic.dev)

### B. 版本历史

| 版本 | 日期 | 变更说明 |
|------|------|---------|
| 1.0.0 | 2026-01-03 | 初始版本，架构设计 |
| 1.1.0 | 2026-01-04 | 完成 MCP Tools 设计，确定 27 个工具列表和 Schema 设计 |

### C. 待办事项

- [ ] 实现阶段 1：基础设施
- [ ] 实现阶段 2：Word Tools（第一批）
- [ ] 实现阶段 3：PowerPoint Tools（第二批）
- [ ] 实现阶段 4：Excel Tools（第三批）
- [ ] 实现阶段 5：完善和测试
- [ ] 编写使用文档
- [ ] 性能优化
