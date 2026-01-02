<!--
文件名: office4ai_dev_plan.md
作者: JQQ
创建日期: 2025/12/18
最后修改日期: 2025/12/18
版权: 2023 JQQ. All rights reserved.
依赖: None
描述: Office4AI 工程开发方案（参考 /arch + examples/ide4ai）
-->

# Office4AI 工程开发方案（参考 `/arch` + `examples/ide4ai`）

## 1. 项目目标（对齐 `/arch`）
- **核心目标**：基于 MCP 协议实现一个可“管理并编辑 Office 文档”的 MCP Server（`office4ai-mcp`）。
- **对标基线**：架构与 `examples/ide4ai` 相似（同质：文件/Workspace 的增删改查 + 工具/资源体系）。
- **差异化重点**：
  - **多模态内容**（文档结构、图片、表格、图表、样式）优先于“严格语法结构校验”。
  - **不做大而全的额外工具封装**（如 ide4ai 的 Terminal），浏览器/Playwright 等能力作为独立 MCP/独立工具组合接入即可。
  - **工程落点**：提供一套“可组合工具链”，让上层 Agent 能完成“检索资料→组织内容→编辑文档→导出/校验”的闭环。

## 2. 参考实现抽象（从 `examples/ide4ai` 提炼可复用模式）
`examples/ide4ai` 的可迁移架构要点（强烈建议 office4ai 沿用）：
- **BaseMCPServer**：统一实现 MCP handlers（`list_tools` / `call_tool` / `list_resources` / `read_resource`）以及 transport（stdio/sse/streamable-http）。
- **Tool/Resource 基类**：
  - `BaseTool`：`name/description/input_schema/execute`，并提供 `validate_input`（Pydantic）。
  - `BaseResource`：`uri/base_uri/name/description/mime_type/read/update_from_uri`；用 `base_uri` 做 key 支持“同资源不同 query 参数”。
- **工程组织方式**：
  - `tools/__init__.py` 聚合工具
  - `resources/__init__.py` 聚合资源
  - `schemas`/`dtos` 承担输入输出结构化（Pydantic）
  - `server.py` 作为 MCP 入口（并在子类中 `_register_tools/_register_resources`）

你当前仓库里 `office4ai/a2c_smcp/server.py` 已经具备类似 `BaseMCPServer`，`office4ai/office/mcp/server.py` 已经有 `OfficeMCPServer` 壳子（但 `_register_tools/_register_resources` 还没实现），非常适合按 ide4ai 模式继续往下填充。

## 3. 目标架构设计（Office4AI 的模块拆分与职责）
以 `office4ai/` 包为核心，建议划分为 5 个层次：

### 3.1 MCP 接入层（Server Layer）
- **模块**：`office4ai/office/mcp/`
- **职责**：
  - 定义 `OfficeMCPServer(BaseMCPServer)`，完成 `_register_tools/_register_resources`
  - 提供 `async_main()` 启动入口（stdio/sse/streamable-http）
- **关键约束**：
  - MCP 层不直接做重业务，只负责：注册、参数校验、错误转换、返回格式标准化。

### 3.2 领域能力层（Office Domain Layer）
- **模块**：`office4ai/office/` + `office4ai/uno_bridge/`
- **职责**：
  - 文档打开/创建/保存/另存为/导出（PDF/图片/HTML）
  - 文档结构操作（段落/标题/列表/表格/图片/页眉页脚/样式）
  - 可选：变更追踪（操作日志、回滚能力、差异摘要）
- **建议抽象**：
  - `DocumentSession`：封装一个“当前文档上下文”（路径、类型、句柄、状态、最近一次渲染快照）
  - `OfficeAppController`：封装 LibreOffice/UNO 进程生命周期、连接、健康检查与重连

### 3.3 工具层（Tools Layer）
- **模块**：`office4ai/a2c_smcp/tools/` + `office4ai/office/tools/`（二选一，建议统一放在 `office4ai/office/tools` 更清晰）
- **职责**：向 MCP 暴露“最小但可组合”的动作单元。
- **工具分组建议**（按“文档工作流”组织）：
  1. **Workspace/文件类**（对齐 ide4ai 的 read/write/edit/glob/grep 思路，但更面向内容资产）
     - `ListWorkspaceFiles` / `ReadAsset` / `WriteAsset`（用于 doc 模板、图片素材、引用资料）
  2. **Document 生命周期类**
     - `DocOpen` / `DocCreate` / `DocSave` / `DocSaveAs` / `DocClose`
  3. **结构化编辑类（核心）**
     - `DocInsertText`（指定锚点：光标/书签/标题/段落索引）
     - `DocReplaceText`（支持范围、正则可选、大小写）
     - `DocApplyStyle`（段落/字符样式）
     - `DocInsertImage`（本地图片路径或 bytes/base64）
     - `DocInsertTable` / `DocUpdateTable`
     - `DocInsertChart`（可先做占位，后续迭代）
  4. **渲染与多模态读取类（office4ai 的关键差异）**
     - `DocRenderPagesToImages`（用于给模型“看”文档效果）
     - `DocExtractOutline`（导出大纲、标题层级）
     - `DocExportPDF`
  5. **检索协作类（与 Playwright/外部 MCP 协作）**
     - office4ai 自身不封装 Playwright MCP，但需要提供“接入点约定”：比如把检索到的资料落地为资产文件，再通过 `ReadAsset`/`Insert` 工具引用。

### 3.4 资源层（Resources Layer）
- **模块**：`office4ai/a2c_smcp/resources/` + `office4ai/office/resources/`
- **职责**：以 MCP resource 方式暴露“当前状态/快照”给 Agent 读取。
- **资源建议**：
  - **`office://document/active`**：当前文档的元信息（路径、类型、页数、上次保存时间、dirty 状态）
  - **`office://document/outline`**：标题大纲（JSON）
  - **`office://document/page?index=1`**：某页渲染图（`image/png` 或 base64 包装）
  - **`office://workspace/assets`**：素材列表（图片、引用文件）

### 3.5 DTO/Schema 层（Contracts Layer）
- **模块**：`office4ai/dtos/` 或 `office4ai/a2c_smcp/schemas.py`
- **职责**：
  - 所有工具入参/出参 Pydantic 模型
  - 错误结构（`error_code`、`error_message`、`details`）
  - 统一返回 `success/message/metadata`（对齐 ide4ai 的输出风格，降低上层适配成本）

## 4. 里程碑规划（工程团队可按周推进）
下面按“可交付物 + 验收标准”拆解，确保每一阶段都能跑通闭环。

### Milestone 0：工程基线与可运行 MCP Server
- **交付物**：
  - `OfficeMCPServer` 可启动（stdio + sse + http三者均兼容）
  - `list_tools/list_resources` 可返回空列表但不报错
  - 统一日志、配置读取（`MCPServerConfig`）
- **验收标准**：
  - `uv run ...` 启动成功
  - MCP 客户端能连接并完成 handshake

### Milestone 1：文档会话与基础生命周期工具
- **交付物**：
  - UNO/LibreOffice 进程拉起、连接、关闭（最小稳定）
  - 工具：`DocOpen/DocCreate/DocSave/DocClose`
  - 资源：`office://document/active`
- **验收标准**：
  - 打开/新建/保存/关闭可重复执行 N 次无崩溃
  - 异常（文件不存在/格式不支持）有清晰错误返回

### Milestone 2：结构化编辑 MVP
- **交付物**：
  - 工具：`DocInsertText`、`DocReplaceText`、`DocApplyStyle`
  - 支持“按标题定位/按段落索引定位/全文替换”至少两种锚点方式
  - 导出：`DocExportPDF`
- **验收标准**：
  - 给定一个模板文档，能自动生成一份包含标题层级与样式的文档并导出 PDF
  - 替换操作可控且可复现（同输入得到同结果）

### Milestone 3：多模态能力（渲染/图片/表格）
- **交付物**：
  - 工具：`DocInsertImage`、`DocInsertTable/DocUpdateTable`
  - 资源：`office://document/page?index=`（渲染输出）
- **验收标准**：
  - 插入图片后，渲染页图能反映变化
  - 表格更新稳定（行列增删改）

### Milestone 4：资产与检索协作闭环
- **交付物**：
  - `WriteAsset/ReadAsset/ListAssets`
  - “外部检索→写入 asset→插入文档”示例流程（放到 `examples/`）
- **验收标准**：
  - 一条脚本/用例能完成：检索素材（由外部完成）→落地→插入→导出

### Milestone 5：稳定性与发布（持续迭代）
- **交付物**：
  - 回归测试、并发策略、超时与重连
  - `pyproject.toml` 版本发布流程（CI）
- **验收标准**：
  - 连续运行 30-60 分钟、执行高频编辑无资源泄漏
  - 集成测试在 CI 通过（可用 mock 或条件跳过需要 GUI/LibreOffice 的部分）

## 5. 团队分工建议（角色与责任边界）
- **[MCP/平台工程]**
  - 负责 `office4ai/a2c_smcp/*`、`OfficeMCPServer` 注册、协议兼容、传输模式、错误返回标准化
- **[Office/UNO 工程]**
  - 负责 `uno_bridge`、LibreOffice 进程管理、文档句柄、格式兼容、渲染导出
- **[工具产品化/体验]**
  - 负责工具拆分粒度、输入输出 schema 设计、示例工作流、资源定义（page/outline）
- **[测试与质量]**
  - 负责集成测试框架、测试文档/素材、稳定性与性能基线、CI 落地

## 6. 工程规范（强制约束，保证团队协作效率）

### 6.1 工具与资源的“协议级规范”
- **工具返回结构统一**（建议）：
  - `success: bool`
  - `message: str`（面向人类的说明）
  - `data: {...}`（结构化结果）
  - `error: {code, message, details}`（失败时）
- **资源 `uri` 规范**：
  - `office://document/...`
  - `office://workspace/...`
  - 需要 query 参数的资源必须实现 `update_from_uri`，并以 `base_uri` 注册。

### 6.2 错误处理与日志
- **错误码分层**：
  - `MCP_*`（协议/注册层）
  - `DOC_*`（文档操作）
  - `UNO_*`（UNO 通讯/进程）
  - `ASSET_*`（素材与文件）
- **日志建议**：
  - Tool 调用：`tool_name`、`doc_id`、`duration_ms`、`success`、`error_code`

### 6.3 测试策略（推荐组合）
- **单元测试**：schema 校验、URI 解析、锚点解析、纯函数转换
- **集成测试**：需要 LibreOffice/UNO 的测试可用 `pytest -m integration` 分组
- **黄金样例（Golden Files）**：
  - 给定输入操作序列，导出 PDF/页图后做 hash 或关键像素区域对比（允许阈值）

### 6.4 运行与调试（对齐当前约束）
- **统一命令**：使用 `uv run` 启动与测试（对齐 `/arch`）
- **开发环境要求**：
  - mac 上 LibreOffice/UNO 可用性检查脚本（建议提供，但不强制第一天做完）

## 7. 风险清单与 POC 建议（先验证再大规模铺开）
- **[UNO 连接稳定性]**
  - 风险：进程挂死、连接断开、并发操作冲突
  - POC：实现 `OfficeAppController.health_check + auto_restart`，压测 100 次 open/edit/export
- **[多模态渲染性能]**
  - 风险：渲染页图耗时、占用 CPU/内存
  - POC：仅渲染“受影响页/指定页”，并缓存上次渲染结果（按文档版本号）
- **[定位锚点的可解释性]**
  - 风险：只靠“第 N 段落”不稳定
  - POC：优先支持“按标题文本/书签/可见文本片段定位”，并提供 `ExtractOutline` 让上层先读再写
- **[格式兼容与导出差异]**
  - 风险：docx/odt/pptx/xlsx 能力差异很大
  - POC：先锁定 1-2 个主格式（例如 `docx + odt`），其它格式走降级策略（只读/只导出）

## 8. 待确认问题（影响一期方案）
- **支持范围**：第一期只做 Writer（文档），还是同时覆盖 Calc/Impress？
- **渲染形态**：资源 page 返回 `image/png`（base64）还是返回落地文件路径更合适？
- **锚点策略**：更倾向“结构锚点（标题/段落索引）”还是“语义锚点（文本片段匹配）”优先？
