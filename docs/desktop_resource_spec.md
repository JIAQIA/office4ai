# Desktop Resource 设计规格

> Office4AI Window 资源实现规格，对齐 A2C-SMCP 协议 Desktop 子系统。

---

## 1. 概述

Desktop 是暴露给 LLM 的 "界面"——通过 `window://` MCP Resource 向 AI Agent 呈现当前 Office 工作区的可视化状态。设计借鉴人类界面思路：激活文档的当前页/当前幻灯片、文档列表、全局元数据摘要等。

### 1.1 设计原则

- **实时优先**：每次 `read()` 通过 Socket.IO 拉取最新数据，3 秒超时后降级
- **聚合视图**：每种文档类型一个 window 路径，内含所有同类文档
- **最后操作即激活**：沿用 `_last_activity` 机制确定激活文档
- **中文渲染**：所有标题、标签使用中文
- **Markdown 结构化**：使用 Markdown 标题和列表组织内容

### 1.2 Phase 总览

| Phase | 交付物 | 核心价值 |
|-------|--------|----------|
| **Phase 1** | WordWindowResource + PptWindowResource + 根索引改造 + 清理 | Agent 可读取实时文档上下文 |
| **Phase 2** | ResourceListChangedNotification (Lifespan + Queue) | Agent 被动感知文档连接变化 |

Phase 1 独立交付即可用（Agent 主动 `read()`），不依赖 Phase 2。

---

## 2. 共享设计（两 Phase 通用）

### 2.1 资源 URI 设计

| 资源 | URI | MIME | 说明 |
|------|-----|------|------|
| 根索引 | `window://office4ai?priority=0&fullscreen=false` | text/plain | 子资源索引总览 |
| Word 窗口 | `window://office4ai/word?priority=50&fullscreen=true` | text/plain | Word 文档聚合视图 |
| PPT 窗口 | `window://office4ai/ppt?priority=50&fullscreen=true` | text/plain | PPT 文档聚合视图 |

### 2.2 URI Query 参数

| 参数 | 类型 | 范围 | 默认值 | 说明 |
|------|------|------|--------|------|
| `priority` | int | [0, 100] | 见上表 | Window 排序优先级 |
| `fullscreen` | bool | true/false | 见上表 | 全屏渲染标志 |
| `range` | int | [0, 10] | 2 | **仅 PPT**: slide 摘要范围 (±N) |

### 2.3 激活文档机制

**激活文档** = 最后一次被 MCP 工具操作的文档。沿用现有 `OfficeWorkspace._last_activity: LastActivity`。

- 每次工具调用后，`BaseTool.execute()` 自动更新 `_last_activity`
- Desktop 渲染时，激活文档在列表中标注 ⭐ 和 `(激活)`
- 仅激活文档展示详细信息（元数据 + 内容/摘要）
- 非激活文档仅在列表中显示 URI
- 无激活文档时（无工具调用过），展示所有文档列表但不展示详情
- WordWindowResource 只关注 `/word` namespace 的激活文档；PptWindowResource 只关注 `/ppt` namespace

### 2.4 类层次

```
BaseResource (已有, ABC)
├── WindowResource (改造为根索引)
├── WordWindowResource (新建)
└── PptWindowResource (新建)
```

### 2.5 文件结构

```
office4ai/a2c_smcp/resources/
├── base.py                    # BaseResource (不变)
├── window.py                  # WindowResource (改造为根索引)
├── word_window.py             # WordWindowResource (新建, Phase 1)
├── ppt_window.py              # PptWindowResource (新建, Phase 1)
└── connected_documents.py     # ❌ 删除 (Phase 1)
```

### 2.6 OASP 协议依赖

| 事件 | 状态 | Phase | 说明 |
|------|------|-------|------|
| `word:get:documentStats` | ✅ 已有 | 1 | 返回页数/字数等 |
| `word:get:visibleContent` | ✅ 已有 | 1 | 返回可见段落 |
| `ppt:get:slideInfo` (无参) | ✅ 已有 | 1 | 返回 slideCount, dimensions, currentSlideIndex |
| `ppt:get:slideInfo` (带 slideIndex) | ✅ 已有 | 1 | 返回单张详情 |

---

# Phase 1: Window 资源渲染

## P1-1. 交付物

| 交付物 | 文件 | 说明 |
|--------|------|------|
| `WordWindowResource` | `resources/word_window.py` | Word 文档聚合窗口资源 |
| `PptWindowResource` | `resources/ppt_window.py` | PPT 文档聚合窗口资源 |
| `WindowResource` 改造 | `resources/window.py` | 简化为根索引 |
| 删除 `ConnectedDocumentsResource` | `resources/connected_documents.py` | 功能已覆盖 |
| 注册更新 | `office/mcp/server.py` | 更新资源注册 |

## P1-2. 渲染格式

### P1-2.1 根索引资源 (`window://office4ai`)

```
# Office 工作区

## 子资源
- window://office4ai/word — Word 文档 (2 个已连接)
- window://office4ai/ppt — PPT 文档 (1 个已连接)
```

无文档连接时：

```
# Office 工作区

## 子资源
暂无文档连接，等待 Office Add-In 接入。
```

### P1-2.2 Word 窗口资源 (`window://office4ai/word`)

```markdown
# Word 工作区

## 文档列表 (2)
- ⭐ file:///Users/jqq/Documents/report.docx (激活)
- file:///Users/jqq/Documents/draft.docx

## 激活文档: report.docx
- 总页数: 12
- 总字数: 3,450
- 段落数: 45

## 当前可见内容
第三章 设计方案

本章节介绍了系统架构设计的核心要点...
3.1 总体架构
采用分层架构，将系统划分为表示层、业务层和数据层...
```

**降级渲染**（实时拉取失败时）：

```markdown
# Word 工作区

## 文档列表 (2)
- ⭐ file:///Users/jqq/Documents/report.docx (激活)
- file:///Users/jqq/Documents/draft.docx

## 激活文档: report.docx
[元数据不可用: 请求超时]

## 当前可见内容
[可见内容不可用: 请求超时]
```

### P1-2.3 PPT 窗口资源 (`window://office4ai/ppt`)

```markdown
# PPT 工作区

## 文档列表 (1)
- ⭐ file:///Users/jqq/Documents/presentation.pptx (激活)

## 激活文档: presentation.pptx
- 总张数: 20
- 尺寸: 960×540 pt (16:9)
- 当前幻灯片: 第 5 张

## 幻灯片摘要 (第 3-7 张)

### 第 3 张: 市场分析
- 元素: 文本框×2, 图表×1, 图片×1
- 备注: 强调Q3增长趋势

### 第 4 张: 竞品对比
- 元素: 文本框×3, 表格×1
- 备注: (无)

### ➡️ 第 5 张: 产品策略 (当前)
- 元素: 文本框×4, 图片×2
- 备注: 重点讲解差异化定位

### 第 6 张: 技术架构
- 元素: 文本框×2, 形状×5
- 备注: 展示微服务拓扑图

### 第 7 张: 里程碑计划
- 元素: 文本框×3, 表格×1
- 备注: (无)
```

**降级渲染**（部分 slide 拉取失败时）：

```markdown
### 第 3 张
[幻灯片信息不可用: 请求超时]
```

## P1-3. 技术实现方案

### P1-3.1 数据获取策略

每次 `read()` 通过 `workspace.emit_to_document()` 向 Add-In 发送 Socket.IO 请求获取最新数据。

| 数据 | Socket.IO 事件 | 说明 |
|------|----------------|------|
| Word 元数据 | `word:get:documentStats` | 页数、字数、段落数 |
| Word 可见内容 | `word:get:visibleContent` | 当前视口可见段落 |
| PPT 元数据 | `ppt:get:slideInfo` (不带 slideIndex) | 总张数、尺寸、当前 slide index |
| PPT 单张摘要 | `ppt:get:slideInfo` (带 slideIndex) | 标题、元素列表、备注 |

**超时与降级**：

- **拉取超时**: 3 秒（独立于工具调用的 30 秒超时）
- **降级策略**: 失败时显示错误提示（如 `[元数据不可用: 请求超时]`），不阻塞其他部分渲染
- **并发优化**: Word 的 stats + visibleContent 并发拉取；PPT 的 ±N 张 slideInfo 用 `asyncio.gather` 并发

**PPT 多张 slide 拉取**：

```
当前 slide index = C (从 ppt:get:slideInfo 无参调用获得)
range = N (默认 2, URI query 可配)
拉取范围 = [max(0, C-N), min(totalSlides-1, C+N)]
并发调用 ppt:get:slideInfo(slideIndex=i) for i in 拉取范围
```

**拉取路径**：

```
WindowResource.read()
  → workspace.emit_to_document(document_uri, event, data)
    → connection_manager.get_socket_by_document()
    → sio.call(event, wrapped_data, to=socket_id)
    → Add-In 返回 ack
```

WindowResource 已持有 `workspace: OfficeWorkspace` 引用，无需引入新抽象。

### P1-3.2 WordWindowResource 实现

```python
class WordWindowResource(BaseResource):
    """window://office4ai/word — Word 文档聚合窗口"""

    FETCH_TIMEOUT = 3  # 秒

    def __init__(self, workspace: OfficeWorkspace, priority: int = 50, fullscreen: bool = True) -> None: ...

    # BaseResource 实现
    uri -> "window://office4ai/word?priority=50&fullscreen=true"
    base_uri -> "window://office4ai/word"
    name -> "Word 工作区"
    mime_type -> "text/plain"

    async def read(self) -> str:
        # 1. 获取 /word namespace 的所有连接文档
        # 2. 确定激活文档 (last_activity 且 namespace == /word)
        # 3. 并发拉取: documentStats + visibleContent (3s 超时)
        # 4. 渲染 Markdown

    async def _fetch_word_stats(self, document_uri: str) -> dict | None:
        """word:get:documentStats, 3s 超时返回 None"""

    async def _fetch_visible_content(self, document_uri: str) -> str | None:
        """word:get:visibleContent, 3s 超时返回 None"""
```

### P1-3.3 PptWindowResource 实现

```python
class PptWindowResource(BaseResource):
    """window://office4ai/ppt — PPT 文档聚合窗口"""

    FETCH_TIMEOUT = 3  # 秒
    DEFAULT_RANGE = 2  # ±N slides

    def __init__(self, workspace: OfficeWorkspace, priority: int = 50, fullscreen: bool = True) -> None: ...

    # BaseResource 实现
    uri -> "window://office4ai/ppt?priority=50&fullscreen=true"
    base_uri -> "window://office4ai/ppt"
    name -> "PPT 工作区"
    mime_type -> "text/plain"

    _range: int = 2  # 可通过 ?range=N 配置

    async def read(self) -> str:
        # 1. 获取 /ppt namespace 的所有连接文档
        # 2. 确定激活文档
        # 3. 拉取 ppt:get:slideInfo (无参) 获取元数据 + currentSlideIndex
        # 4. 并发拉取 ±N 张 ppt:get:slideInfo(slideIndex=i)
        # 5. 渲染 Markdown

    async def _fetch_presentation_info(self, document_uri: str) -> dict | None:
        """ppt:get:slideInfo 不带 slideIndex, 3s 超时"""

    async def _fetch_slide_summaries(
        self, document_uri: str, center: int, range_n: int, total: int
    ) -> list[dict | None]:
        """并发拉取 ±N 张 slideInfo, 每张独立 3s 超时"""
```

### P1-3.4 WindowResource 改造（根索引）

```python
class WindowResource(BaseResource):
    """window://office4ai — 根索引资源"""

    def __init__(self, workspace: OfficeWorkspace, priority: int = 0, fullscreen: bool = False) -> None: ...

    uri -> "window://office4ai?priority=0&fullscreen=false"
    base_uri -> "window://office4ai"
    name -> "Office 工作区"

    async def read(self) -> str:
        # 1. 统计 /word 和 /ppt 各自的连接文档数
        # 2. 渲染子资源索引列表
```

### P1-3.5 注册与集成

```python
# OfficeMCPServer._register_resources()
def _register_resources(self):
    root = WindowResource(self.workspace, priority=0, fullscreen=False)
    word = WordWindowResource(self.workspace, priority=50, fullscreen=True)
    ppt = PptWindowResource(self.workspace, priority=50, fullscreen=True)

    self.resources[root.base_uri] = root
    self.resources[word.base_uri] = word
    self.resources[ppt.base_uri] = ppt
```

`BaseMCPServer.read_resource()` 路由无需调整。当前实现 `f"{parsed.scheme}://{parsed.netloc}{parsed.path}"` 已正确处理子路径。

### P1-3.6 迁移与清理

- 删除 `office4ai/a2c_smcp/resources/connected_documents.py`
- 从 `OfficeMCPServer._register_resources()` 中移除注册
- 从 `resources/__init__.py` 中移除导出

## P1-4. 测试覆盖方案

### P1-4.1 测试矩阵

| 层级 | 测试文件 | 标记 | 覆盖目标 |
|------|----------|------|----------|
| **单元** | `tests/unit_tests/.../resources/test_word_window_resource.py` | (无) | 渲染逻辑、降级、筛选 |
| **单元** | `tests/unit_tests/.../resources/test_ppt_window_resource.py` | (无) | 渲染逻辑、范围计算、降级 |
| **单元** | `tests/unit_tests/.../resources/test_window_resource.py` | (无) | 根索引渲染（更新现有） |
| **契约** | `tests/contract_tests/resources/test_word_window_contract.py` | `@contract` | 真实 Socket.IO + MockAddInClient |
| **契约** | `tests/contract_tests/resources/test_ppt_window_contract.py` | `@contract` | 真实 Socket.IO + MockAddInClient |
| **集成** | `tests/integration_tests/.../mcp/test_mcp_protocol.py` | `@integration` | MCP list_resources + read_resource |

### P1-4.2 单元测试详细设计

**Mock 策略**: Mock `workspace.emit_to_document()` 和 `connection_manager.get_all_clients()`

#### test_word_window_resource.py

```python
class TestWordWindowResource:
    """WordWindowResource 单元测试"""

    # --- 基础属性 ---
    def test_uri_format(self): ...              # URI 包含 priority 和 fullscreen
    def test_base_uri(self): ...                # "window://office4ai/word"
    def test_name(self): ...                    # "Word 工作区"
    def test_update_from_uri_priority(self): ...  # query 参数解析

    # --- 渲染：无文档 ---
    async def test_read_no_documents(self): ...  # 空列表渲染

    # --- 渲染：有文档无激活 ---
    async def test_read_documents_no_active(self):
        """有连接文档但无 last_activity → 列表但无详情"""

    # --- 渲染：正常激活 ---
    async def test_read_with_active_document(self):
        """Mock emit_to_document 返回 stats + visibleContent → 完整 Markdown"""
        # 验证: 标题、文档列表 (⭐ 标注)、元数据、可见内容

    # --- 渲染：多文档聚合 ---
    async def test_read_multiple_word_documents(self):
        """2 个 Word 文档连接 → 列表显示 2 个, 仅激活文档有详情"""

    # --- Namespace 筛选 ---
    async def test_ignores_ppt_documents(self):
        """连接中混合 /word 和 /ppt → 只渲染 /word 文档"""

    async def test_ignores_ppt_active_document(self):
        """last_activity 指向 /ppt 文档 → Word 视图无激活文档"""

    # --- 降级：stats 超时 ---
    async def test_read_stats_timeout(self):
        """emit_to_document(documentStats) 抛 TimeoutError → 显示错误提示"""

    # --- 降级：visibleContent 超时 ---
    async def test_read_visible_content_timeout(self):
        """emit_to_document(visibleContent) 抛 TimeoutError → 显示错误提示"""

    # --- 降级：全部超时 ---
    async def test_read_all_fetch_timeout(self):
        """两个拉取都超时 → 文档列表正常, 详情全部降级"""

    # --- 降级：部分成功 ---
    async def test_read_stats_ok_content_timeout(self):
        """stats 成功 visibleContent 超时 → 元数据正常, 内容降级"""

    # --- 并发验证 ---
    async def test_concurrent_fetch(self):
        """验证 stats 和 visibleContent 是并发拉取 (通过 mock 计时验证)"""
```

#### test_ppt_window_resource.py

```python
class TestPptWindowResource:
    """PptWindowResource 单元测试"""

    # --- 基础属性 ---
    def test_uri_format(self): ...
    def test_update_from_uri_range(self): ...    # ?range=3 → self._range = 3

    # --- 渲染：正常 ---
    async def test_read_with_active_presentation(self):
        """Mock 返回 20 slides, currentSlideIndex=4 → 显示第 3-7 张摘要"""
        # 验证: 元数据 (张数/尺寸/当前), 摘要列表, ➡️ 标注当前

    # --- 范围计算 ---
    async def test_slide_range_at_beginning(self):
        """currentSlideIndex=0, range=2 → 显示第 0-2 张"""

    async def test_slide_range_at_end(self):
        """currentSlideIndex=19, totalSlides=20, range=2 → 显示第 17-19 张"""

    async def test_slide_range_small_presentation(self):
        """totalSlides=3, range=2 → 显示全部 3 张"""

    # --- 摘要字段 ---
    async def test_slide_summary_fields(self):
        """验证每张 slide 渲染: 标题、元素计数 (按类型)、备注"""

    async def test_slide_summary_no_notes(self):
        """备注为空 → 显示 '(无)'"""

    # --- 降级 ---
    async def test_presentation_info_timeout(self):
        """ppt:get:slideInfo 无参超时 → 元数据区域降级"""

    async def test_partial_slide_timeout(self):
        """5 张中 2 张超时 → 超时的显示错误提示, 其余正常"""

    async def test_all_slides_timeout(self):
        """全部 slide 拉取超时 → 元数据正常 (假设无参调用成功), 摘要全部降级"""

    # --- 并发验证 ---
    async def test_concurrent_slide_fetch(self):
        """验证 ±N 张 slideInfo 是并发拉取"""

    # --- 多文档 / Namespace ---
    async def test_ignores_word_documents(self): ...
    async def test_multiple_ppt_documents(self): ...
```

#### test_window_resource.py (更新现有)

```python
class TestWindowResourceIndex:
    """WindowResource 根索引 单元测试"""

    async def test_read_empty(self): ...           # 无连接 → "暂无文档连接"
    async def test_read_word_only(self): ...        # 2 个 Word → word (2), ppt (0)
    async def test_read_ppt_only(self): ...         # 1 个 PPT → word (0), ppt (1)
    async def test_read_mixed(self): ...            # Word + PPT 混合
    async def test_read_dedup_by_uri(self): ...     # 同一文档多个连接 → 去重
    def test_uri_format_changed(self): ...          # fullscreen=false
    def test_priority_changed(self): ...            # priority=0
```

### P1-4.3 契约测试详细设计

**策略**: 复用 `contract_tests` 现有基础设施（`MockAddInClient` + `factories` + 真实 Socket.IO server on port 3003 + 真实 `OfficeWorkspace`）。验证 WindowResource 通过 workspace 发出的 Socket.IO 请求格式正确、能正确解析 Add-In 响应并渲染。

#### conftest.py 新增 fixtures

```python
# tests/contract_tests/resources/conftest.py

@pytest_asyncio.fixture
async def word_window_resource(workspace: OfficeWorkspace) -> WordWindowResource:
    return WordWindowResource(workspace, priority=50, fullscreen=True)

@pytest_asyncio.fixture
async def ppt_window_resource(workspace: OfficeWorkspace) -> PptWindowResource:
    return PptWindowResource(workspace, priority=50, fullscreen=True)
```

#### test_word_window_contract.py

```python
@pytest.mark.asyncio
@pytest.mark.contract
class TestWordWindowContract:
    """WordWindowResource 契约测试 — 真实 Socket.IO, MockAddInClient 响应"""

    async def test_read_fetches_stats_and_content(
        self,
        word_window_resource: WordWindowResource,
        mock_word_client_factory,
        word_factory,
    ):
        """
        完整链路验证:
        1. MockAddInClient 连接到 /word namespace
        2. 注册 documentStats 和 visibleContent 的响应
        3. 调用工具设置 last_activity (使文档成为激活文档)
        4. word_window_resource.read() 触发 Socket.IO 请求
        5. 验证渲染结果包含 stats 和 content
        """
        # Setup: 创建 mock client 并注册响应
        client = mock_word_client_factory(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_word_client",
            document_uri="file:///tmp/contract_test.docx",
        )
        client.register_response("word:get:documentStats", lambda req: {
            "requestId": req["requestId"],
            "success": True,
            "data": {"pageCount": 5, "wordCount": 1200, "paragraphCount": 20},
            "timestamp": ..., "duration": 10,
        })
        client.register_response("word:get:visibleContent", lambda req: {
            "requestId": req["requestId"],
            "success": True,
            "data": {"content": "Hello World\nParagraph 2"},
            "timestamp": ..., "duration": 10,
        })
        await client.connect()

        # 设置 last_activity 使此文档成为激活文档
        word_window_resource.workspace.update_last_activity(
            client.document_uri, "word_get_visible_content", {}
        )

        # Act
        content = await word_window_resource.read()

        # Assert: 渲染结果包含 stats 和 content
        assert "文档列表 (1)" in content
        assert "⭐" in content
        assert "总页数: 5" in content
        assert "总字数: 1,200" in content  # 或 1200
        assert "Hello World" in content

        await client.disconnect()

    async def test_read_timeout_graceful_degradation(
        self,
        word_window_resource: WordWindowResource,
        mock_word_client_factory,
    ):
        """
        超时降级验证:
        1. MockAddInClient 注册延迟 5s 的响应 (超过 3s 超时)
        2. read() 应在 ~3s 后返回降级渲染
        3. 不应抛出异常
        """
        client = mock_word_client_factory(...)

        async def slow_response(req):
            await asyncio.sleep(5)  # 超过 3s 超时
            return {...}

        client.register_response("word:get:documentStats", slow_response)
        client.register_response("word:get:visibleContent", slow_response)
        await client.connect()

        word_window_resource.workspace.update_last_activity(
            client.document_uri, "word_get_visible_content", {}
        )

        content = await word_window_resource.read()

        # 降级渲染: 文档列表正常, 详情显示错误
        assert "文档列表 (1)" in content
        assert "不可用" in content or "超时" in content

        await client.disconnect()

    async def test_read_no_active_document(
        self,
        word_window_resource: WordWindowResource,
        mock_word_client_factory,
    ):
        """无 last_activity → 列表正常, 无详情, 不发送 Socket.IO 请求"""
        client = mock_word_client_factory(...)
        await client.connect()

        content = await word_window_resource.read()

        assert "文档列表 (1)" in content
        assert "激活文档" not in content
        # 验证没有 Socket.IO 请求发出 (client.received_events 为空)
        assert len(client.received_events) == 0

        await client.disconnect()
```

#### test_ppt_window_contract.py

```python
@pytest.mark.asyncio
@pytest.mark.contract
class TestPptWindowContract:
    """PptWindowResource 契约测试"""

    async def test_read_fetches_presentation_and_slides(
        self,
        ppt_window_resource: PptWindowResource,
        mock_word_client_factory,  # 复用 factory, namespace=/ppt
        ppt_factory,
    ):
        """
        完整链路验证:
        1. MockAddInClient 连接到 /ppt namespace
        2. 注册 slideInfo 无参和带参的响应
        3. ppt_window_resource.read() → 1 次无参 + N 次带参并发调用
        4. 验证渲染包含元数据 + slide 摘要
        """
        client = mock_word_client_factory(
            server_url="http://127.0.0.1:3003",
            namespace="/ppt",
            client_id="contract_ppt_client",
            document_uri="file:///tmp/contract_test.pptx",
        )

        # 无参调用返回 presentation 元数据
        def slideinfo_response(req):
            slide_index = req.get("slideIndex")
            if slide_index is None:
                # 无参: 返回 presentation 级别信息
                return {
                    "requestId": req["requestId"],
                    "success": True,
                    "data": {
                        "slideCount": 10,
                        "currentSlideIndex": 3,
                        "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
                    },
                    "timestamp": ..., "duration": 10,
                }
            else:
                # 带参: 返回单张 slide 详情
                return {
                    "requestId": req["requestId"],
                    "success": True,
                    "data": ppt_factory.slide_info_response(slide_index=slide_index),
                    "timestamp": ..., "duration": 10,
                }

        client.register_response("ppt:get:slideInfo", slideinfo_response)
        await client.connect()

        ppt_window_resource.workspace.update_last_activity(
            client.document_uri, "ppt_get_slide_info", {}
        )

        content = await ppt_window_resource.read()

        # Assert: 元数据
        assert "总张数: 10" in content
        assert "当前幻灯片: 第 4 张" in content  # 0-based → 1-based
        # Assert: 摘要范围 (slide 1-5, 即 index 1-5)
        assert "➡️" in content  # 当前 slide 标注
        # Assert: 摘要字段 (标题/元素/备注)
        assert "元素:" in content

        await client.disconnect()

    async def test_concurrent_slide_fetch_count(
        self,
        ppt_window_resource: PptWindowResource,
        mock_word_client_factory,
    ):
        """验证并发拉取次数: 1 次无参 + min(2*N+1, total) 次带参"""
        # Track call count via response factory
        call_count = {"value": 0}

        def counting_response(req):
            call_count["value"] += 1
            return {...}

        # ... setup, read, assert call_count matches expected

    async def test_partial_slide_timeout(self, ...):
        """5 张中 2 张超时 → 超时的显示错误, 其余正常"""
```

### P1-4.4 集成测试详细设计

**策略**: 扩展现有 `test_mcp_protocol.py`，验证 MCP 协议层面资源注册和读取正确。

#### test_mcp_protocol.py 新增用例

```python
# 追加到现有 TestMCPProtocol 类

@pytest.mark.integration
class TestMCPResourcesPhase1:
    """Phase 1: MCP 资源层面集成测试"""

    async def test_list_resources_includes_window_resources(self):
        """
        通过 MCP stdio 协议验证:
        1. list_resources() 返回 3 个资源 (root + word + ppt)
        2. 不再包含 office://workspace/documents
        3. URI 格式正确
        """
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                resources = await session.list_resources()
                uris = [str(r.uri) for r in resources.resources]

                # 新资源存在
                assert any("window://office4ai/word" in u for u in uris)
                assert any("window://office4ai/ppt" in u for u in uris)
                assert any(u.startswith("window://office4ai?") for u in uris)

                # 旧资源已删除
                assert not any("office://workspace/documents" in u for u in uris)

    async def test_read_window_root_resource(self):
        """
        读取根索引资源:
        1. 无 Add-In 连接时 → 返回 "暂无文档连接"
        """
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.read_resource("window://office4ai")
                content = result.contents[0].text
                assert "Office 工作区" in content
                assert "暂无文档连接" in content

    async def test_read_word_window_no_connection(self):
        """
        读取 Word 窗口资源（无 Add-In 连接）:
        1. 返回空文档列表
        2. 不抛异常
        """
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.read_resource("window://office4ai/word")
                content = result.contents[0].text
                assert "Word 工作区" in content
                assert "文档列表 (0)" in content

    async def test_read_ppt_window_no_connection(self):
        """读取 PPT 窗口资源（无 Add-In 连接）"""
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.read_resource("window://office4ai/ppt")
                content = result.contents[0].text
                assert "PPT 工作区" in content

    async def test_resource_count_updated(self):
        """list_resources 返回正确的资源数量 (Phase 1 后应为 3)"""
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                resources = await session.list_resources()
                assert len(resources.resources) == 3
```

### P1-4.5 测试覆盖总结

| 测试层级 | 用例数 | 核心验证点 |
|----------|--------|------------|
| **单元测试** | ~30 | 渲染格式、降级逻辑、范围计算、namespace 筛选、URI 参数解析 |
| **契约测试** | ~8 | 真实 Socket.IO 通信、请求格式、响应解析、超时降级、并发拉取 |
| **集成测试** | ~5 | MCP 协议层资源注册、URI 路由、list/read 端到端 |

---

# Phase 2: ResourceListChanged 通知机制

## P2-1. 交付物

| 交付物 | 文件 | 说明 |
|--------|------|------|
| Lifespan 函数 | `a2c_smcp/server.py` | MCP Server lifespan 共享状态 |
| 通知消费者 | `a2c_smcp/server.py` | 后台 task 消费 queue 发送通知 |
| connect 回调 | `socketio/services/connection_manager.py` | 新增 `register_connect_callback` |
| MCP Capabilities | `a2c_smcp/server.py` | 声明 `resources.listChanged: true` |
| OfficeMCPServer 集成 | `office/mcp/server.py` | 注册回调、启动消费者 |

## P2-2. 触发时机

| 事件 | 通知 |
|------|------|
| Add-In 连接 (on_connect) | `notifications/resources/list_changed` |
| Add-In 断连 (on_disconnect) | `notifications/resources/list_changed` |

## P2-3. 技术实现方案

### P2-3.1 架构概览

```
connection_manager on_connect/on_disconnect 回调
  → notification_queue.put_nowait("resource_list_changed")

MCP Server 后台 task (_notification_consumer):
  while True:
    event = await notification_queue.get()
    await session.send_resource_list_changed()
```

采用 MCP SDK 官方 `lifespan` 机制实现跨上下文通知，通过 Queue 解耦 Socket.IO 上下文和 MCP handler 上下文。

### P2-3.2 connection_manager 扩展

当前只有 `register_disconnect_callback`，需新增 `register_connect_callback`：

```python
class ConnectionManager:
    def __init__(self):
        ...
        self._connect_callbacks: list[Callable[[str], None]] = []

    def register_connect_callback(self, callback: Callable[[str], None]) -> None:
        """注册连接回调, document_uri 作为参数"""
        self._connect_callbacks.append(callback)

    def register_client(self, ...):
        ...
        # 注册完成后触发回调
        for callback in self._connect_callbacks:
            try:
                callback(document_uri)
            except Exception as e:
                logger.warning(f"Connect callback error: {e}")
```

### P2-3.3 Lifespan 函数

```python
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

@asynccontextmanager
async def office_mcp_lifespan(server: Server) -> AsyncIterator[dict]:
    """MCP Server lifespan — 创建共享通知基础设施"""
    state: dict = {
        "session": None,           # 首次 handler 调用时捕获
        "notification_queue": asyncio.Queue(),
    }
    yield state
```

### P2-3.4 Session 捕获

在 `_setup_handlers` 的 handler 中首次捕获 session：

```python
def _setup_handlers(self) -> None:
    @self.server.list_resources()
    async def list_resources():
        # 首次调用时捕获 session
        ctx = self.server.request_context
        lifespan = ctx.lifespan_context
        if lifespan.get("session") is None:
            lifespan["session"] = ctx.session
            logger.info("MCP session captured for notifications")
        return [...]

    @self.server.call_tool()
    async def call_tool(name, arguments):
        # 同样在此处捕获 (确保首次调用不论是 list 还是 call 都能捕获)
        ctx = self.server.request_context
        lifespan = ctx.lifespan_context
        if lifespan.get("session") is None:
            lifespan["session"] = ctx.session
        ...
```

### P2-3.5 后台消费 Task

```python
# OfficeMCPServer

async def _async_startup(self) -> None:
    await self.workspace.start()

    # 获取 lifespan state (通过构造时保存的引用)
    self._notification_queue = self._lifespan_state["notification_queue"]

    # 注册连接变更回调
    connection_manager.register_connect_callback(self._on_connection_change)
    connection_manager.register_disconnect_callback(self._on_connection_change)

    # 启动通知消费者
    self._notification_task = asyncio.create_task(self._notification_consumer())

async def _async_shutdown(self) -> None:
    # 停止消费者
    if self._notification_task:
        self._notification_task.cancel()
        try:
            await self._notification_task
        except asyncio.CancelledError:
            pass
    await self.workspace.stop()

def _on_connection_change(self, document_uri: str) -> None:
    """connection_manager 回调, 同步方法"""
    if self._notification_queue:
        self._notification_queue.put_nowait("resource_list_changed")

async def _notification_consumer(self) -> None:
    """后台 task: 从 queue 取事件, 通过 session 发送 MCP 通知"""
    while True:
        event = await self._notification_queue.get()
        session = self._lifespan_state.get("session")
        if session is None:
            logger.debug("MCP session not yet captured, skipping notification")
            continue
        if event == "resource_list_changed":
            try:
                await session.send_resource_list_changed()
                logger.info("Sent ResourceListChangedNotification")
            except Exception as e:
                logger.warning(f"Failed to send ResourceListChanged: {e}")
```

### P2-3.6 MCP Capabilities 声明

```python
# BaseMCPServer.__init__ 或 OfficeMCPServer.__init__
self.server = Server(server_name)
# Server 的 create_initialization_options() 中需要声明:
# capabilities.resources.listChanged = True
```

需要确认 `mcp.Server` 构造函数或 `create_initialization_options()` 如何声明 `listChanged`。如果 SDK 不直接支持，可通过 override `create_initialization_options()` 实现。

## P2-4. 测试覆盖方案

### P2-4.1 测试矩阵

| 层级 | 测试文件 | 标记 | 覆盖目标 |
|------|----------|------|----------|
| **单元** | `tests/unit_tests/.../services/test_connection_manager_callbacks.py` | (无) | connect 回调注册与触发 |
| **单元** | `tests/unit_tests/.../test_notification_consumer.py` | (无) | Queue 消费 + session mock |
| **契约** | `tests/contract_tests/resources/test_notification_contract.py` | `@contract` | 真实 Socket.IO 连接 → Queue 事件 |
| **集成** | `tests/integration_tests/.../mcp/test_mcp_notifications.py` | `@integration` | MCP 协议层通知发送与接收 |

### P2-4.2 单元测试详细设计

#### test_connection_manager_callbacks.py

```python
class TestConnectionManagerConnectCallback:
    """connection_manager connect 回调单元测试"""

    def test_register_connect_callback(self):
        """注册回调后, register_client 时触发"""
        callback = MagicMock()
        connection_manager.register_connect_callback(callback)
        connection_manager.register_client(
            socket_id="s1", client_id="c1",
            document_uri="file:///test.docx", namespace="/word"
        )
        callback.assert_called_once()
        # 参数为 normalized document_uri
        args = callback.call_args[0]
        assert "test.docx" in args[0]

    def test_multiple_connect_callbacks(self):
        """多个回调全部触发"""

    def test_connect_callback_exception_not_propagate(self):
        """回调异常不影响 register_client 正常完成"""
        def bad_callback(uri):
            raise RuntimeError("boom")
        connection_manager.register_connect_callback(bad_callback)
        # 不应抛异常
        connection_manager.register_client(...)

    def test_disconnect_callback_still_works(self):
        """新增 connect 回调不影响现有 disconnect 回调"""
```

#### test_notification_consumer.py

```python
class TestNotificationConsumer:
    """通知消费者 Task 单元测试"""

    async def test_consume_and_send(self):
        """queue 放入事件 → 消费者调用 session.send_resource_list_changed()"""
        mock_session = AsyncMock()
        lifespan_state = {"session": mock_session, "notification_queue": asyncio.Queue()}
        lifespan_state["notification_queue"].put_nowait("resource_list_changed")

        # 启动消费者, 给少许时间消费, 然后取消
        consumer_task = asyncio.create_task(
            mcp_server._notification_consumer()
        )
        await asyncio.sleep(0.1)
        consumer_task.cancel()

        mock_session.send_resource_list_changed.assert_called_once()

    async def test_consume_no_session_skip(self):
        """session 尚未捕获 → 跳过通知, 不崩溃"""
        lifespan_state = {"session": None, "notification_queue": asyncio.Queue()}
        lifespan_state["notification_queue"].put_nowait("resource_list_changed")
        # 消费者应跳过, 不异常

    async def test_consume_send_failure_continue(self):
        """send 异常 → 记录 warning, 继续消费下一个事件"""
        mock_session = AsyncMock()
        mock_session.send_resource_list_changed.side_effect = RuntimeError("broken")
        # 放 2 个事件, 第一个失败不应阻塞第二个

    async def test_consume_unknown_event_ignore(self):
        """未知事件类型 → 忽略"""

    async def test_shutdown_cancellation(self):
        """_async_shutdown 取消 task → 正常退出"""
```

### P2-4.3 契约测试详细设计

**策略**: 真实 Socket.IO 连接/断连 → 验证 Queue 中产生事件。

#### test_notification_contract.py

```python
@pytest.mark.asyncio
@pytest.mark.contract
class TestNotificationContract:
    """通知机制契约测试 — 真实 Socket.IO 触发"""

    async def test_client_connect_triggers_queue_event(
        self,
        workspace: OfficeWorkspace,
        mock_word_client_factory,
    ):
        """
        完整链路 (Socket.IO → Queue):
        1. 注册连接回调绑定到 queue
        2. MockAddInClient 连接
        3. 验证 queue 中有 "resource_list_changed" 事件
        """
        queue = asyncio.Queue()
        connection_manager.register_connect_callback(
            lambda uri: queue.put_nowait("resource_list_changed")
        )

        client = mock_word_client_factory(...)
        await client.connect()
        await asyncio.sleep(0.2)  # 等回调执行

        assert not queue.empty()
        event = queue.get_nowait()
        assert event == "resource_list_changed"

        await client.disconnect()

    async def test_client_disconnect_triggers_queue_event(
        self,
        workspace: OfficeWorkspace,
        mock_word_client_factory,
    ):
        """
        断连链路:
        1. MockAddInClient 连接 → 清空 queue
        2. MockAddInClient 断连
        3. 验证 queue 中有新事件
        """
        queue = asyncio.Queue()
        connection_manager.register_disconnect_callback(
            lambda uri: queue.put_nowait("resource_list_changed")
        )

        client = mock_word_client_factory(...)
        await client.connect()
        # 清空 connect 产生的事件
        while not queue.empty():
            queue.get_nowait()

        await client.disconnect()
        await asyncio.sleep(0.5)  # disconnect 可能有延迟

        assert not queue.empty()

    async def test_multiple_connects_multiple_events(self, ...):
        """3 个 client 连接 → queue 中有 3 个事件"""

    async def test_rapid_connect_disconnect_no_loss(self, ...):
        """快速连接/断连 → 所有事件都进入 queue, 无丢失"""
```

### P2-4.4 集成测试详细设计

**策略**: 通过 MCP stdio 协议验证 Client 端能正确接收到 `ResourceListChangedNotification`。这是最高层级的端到端验证。

#### test_mcp_notifications.py

```python
@pytest.mark.asyncio
@pytest.mark.integration
class TestMCPNotifications:
    """MCP 协议层通知集成测试"""

    async def test_capabilities_declare_list_changed(self):
        """
        MCP 握手后验证 server capabilities 包含 resources.listChanged
        """
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                result = await session.initialize()
                resources_cap = result.capabilities.resources
                assert resources_cap is not None
                assert resources_cap.listChanged is True

    async def test_receive_notification_on_addin_connect(self):
        """
        端到端验证:
        1. 启动 MCP server (含 workspace Socket.IO)
        2. MCP client 连接并初始化
        3. 触发 list_resources 使 session 被捕获
        4. 模拟 Add-In 连接到 Socket.IO
        5. MCP client 应收到 ResourceListChangedNotification
        6. 再次 list_resources 应看到新文档

        注: 此测试需要在 MCP subprocess 的 Socket.IO 端口上
        连接一个真实的 AsyncClient, 验证通知通过 MCP 传递到 client。
        """
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                result = await session.initialize()

                # 首先调用 list_resources 触发 session 捕获
                resources_before = await session.list_resources()

                # 连接 Add-In 到 workspace 的 Socket.IO
                sio_client = socketio.AsyncClient()
                await sio_client.connect(
                    "http://127.0.0.1:3000",
                    namespaces=["/word"],
                    auth={"clientId": "test_addin", "documentUri": "file:///test.docx"},
                )
                await asyncio.sleep(1)  # 等通知处理

                # 再次 list_resources → 验证文档出现
                resources_after = await session.list_resources()
                # 注: 通知本身在 MCP 协议层是 fire-and-forget,
                # 我们通过读取资源内容来间接验证连接已被感知

                # 读取 word window 资源验证新文档出现
                result = await session.read_resource("window://office4ai/word")
                content = result.contents[0].text
                assert "test.docx" in content

                await sio_client.disconnect()

    async def test_notification_not_sent_before_session_capture(self):
        """
        边界条件: session 未捕获前的连接事件不会导致崩溃
        (事件进入 queue 但消费时 session=None → skip)
        """
```

### P2-4.5 测试覆盖总结

| 测试层级 | 用例数 | 核心验证点 |
|----------|--------|------------|
| **单元测试** | ~10 | 回调注册/触发/异常隔离、消费者逻辑、session 未就绪处理 |
| **契约测试** | ~4 | 真实 Socket.IO 连接/断连 → Queue 事件产生 |
| **集成测试** | ~3 | MCP capabilities 声明、端到端通知传递、边界条件 |

---

## 附录 A: 决策记录

| 决策项 | 结论 | 备选方案 |
|--------|------|----------|
| 窗口粒度 | 按类型分路径 `/word`, `/ppt` | 统一单窗口 / 按文档实例分 |
| 多文档处理 | 聚合视图 | 按文档拆路径 |
| 激活文档 | 最后操作即激活 | 显式 activate API / 混合模式 |
| 数据获取路径 | 通过 workspace.emit_to_document() | 直接持有 sio / Provider 抽象 |
| PPT 摘要范围 | ±2 可配置 | 固定 ±1 / 全量精简 / 分页参数化 |
| PPT 摘要字段 | 标题+元素计数+备注 | 仅标题+元素 / 标题+正文摘要 |
| 超时 | 3 秒 | 5 秒 / 10 秒 |
| 渲染语言 | 中文 | 英文 / 双语 |
| 渲染格式 | Markdown 结构化 | 紧凑标记风格 |
| 内容截断 | 原样展示, 交给 Computer 层 | 智能截断 / 可配置 |
| 类设计 | 拆分为独立类 | 统一类+内部分发 / 基类+子类 |
| 根资源 | 保留为索引页 | 删除 / 合并 |
| 老资源清理 | 删除 connected_documents.py | 保留共存 |
| 通知机制 | Lifespan + Queue | 直接存 session / MVP 不做 |
| Word 元数据 | 页数+字数+段落数 | 仅页数+字数 / 全部可用 |

## 附录 B: 参考资料

- [MCP Resources 规范 (2025-06-18)](https://modelcontextprotocol.io/specification/2025-06-18/server/resources)
- [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk)
- [FastMCP Context 文档](https://gofastmcp.com/python-sdk/fastmcp-server-context)
- [A2C-SMCP Desktop 规范](../../../A2C-SMCP/a2c-smcp-protocol/docs/specification/desktop.md)
- [A2C-SMCP Events 规范](../../../A2C-SMCP/a2c-smcp-protocol/docs/specification/events.md)

---

**最后更新**: 2026-03-10
**维护者**: JQQ <jqq1716@gmail.com>
