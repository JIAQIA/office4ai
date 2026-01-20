# word:select:text E2E Tests

`word:select:text` 事件的手动端到端测试。

## 📝 参数命名规范

**重要**：所有测试代码使用 **snake_case**（符合 Python PEP 8 规范），DTO 系统会自动转换为协议层的 **camelCase**。

```python
# ✅ 正确：使用 snake_case
params = {
    "document_uri": document_uri,
    "search_text": search_text,
    "selection_mode": selection_mode,
    "select_index": select_index,
    "search_options": {"matchCase": True}
}

# ❌ 错误：不要直接使用 camelCase
params = {
    "document_uri": document_uri,  # snake_case
    "searchText": search_text,      # ❌ 应该是 search_text
}
```

**转换机制**：
- Python DTO 定义：`search_text: str = Field(..., alias="searchText")`
- 发送到 Add-In 时自动转换为：`{"searchText": "..."}`
- Pydantic 使用 `by_alias=True` 序列化

**相关文档**：
- DTO 定义：`office4ai/environment/workspace/dtos/word.py`
- 协议规范：[Confluence - word:select:text](https://turingfocus.atlassian.net/wiki/pages/42467331)

---

## 测试场景总览

| 编号 | 测试名称 | 参数 | 实现状态 |
|-----|---------|------|---------|
| 1 | 简单选中文本 | `{searchText: "Hello World"}` | ✅ `test_basic_select.py#1` |
| 2 | 选择第N个匹配项 | `{searchText: "test", selectIndex: 2}` | ✅ `test_basic_select.py#2` |
| 3 | 不区分大小写 | `{searchText: "hello", matchCase: false}` | ✅ `test_basic_select.py#3` |
| 4 | 全字匹配 | `{searchText: "test", matchWholeWord: true}` | ✅ `test_basic_select.py#4` |
| 5 | 通配符搜索 | `{searchText: "test*", matchWildcards: true}` | ✅ `test_search_options.py#4` |
| 6 | select 模式 | `{selectionMode: "select"}` | ✅ `test_selection_modes.py#1` |
| 7 | start 模式 | `{selectionMode: "start"}` | ✅ `test_selection_modes.py#2` |
| 8 | end 模式 | `{selectionMode: "end"}` | ✅ `test_selection_modes.py#3` |
| 9 | 未找到匹配 | `{searchText: "nonexistent"}` | ✅ `test_edge_cases.py#1` |
| 10 | 空搜索文本 | `{searchText: ""}` | ✅ `test_edge_cases.py#2` |
| 11 | 特殊字符搜索 | `{searchText: "@#$%"}` | ✅ `test_edge_cases.py#4` |
| 12 | 组合搜索选项 | `{matchCase: true, matchWholeWord: true}` | ✅ `test_search_options.py#5` |
| 13 | 超出索引范围 | `{selectIndex: 999}` | ✅ `test_edge_cases.py#3` |
| 14 | 中文字符搜索 | `{searchText: "中文"}` | ❌ 未实现 |
| 15 | 长文本搜索 | `{searchText: "very long text..."}` | ✅ `test_edge_cases.py#5` |
| 16 | 数字搜索 | `{searchText: "12345"}` | ❌ 未实现 |

**已实现：14/16 (87.5%)**

## 运行方式

### 前置条件
1. 启动 Workspace 服务器（测试会自动启动）
2. 在 Word 中加载 office-editor4ai Add-In
3. 确保文档已打开并连接

### 运行测试

```bash
# 基础选择测试（4 个）
uv run python manual_tests/select_text_e2e/test_basic_select.py --test all

# 搜索选项测试（5 个）
uv run python manual_tests/select_text_e2e/test_search_options.py --test all

# 选择模式测试（4 个）
uv run python manual_tests/select_text_e2e/test_selection_modes.py --test all

# 边界情况测试（5 个）
uv run python manual_tests/select_text_e2e/test_edge_cases.py --test all

# 运行单个测试
uv run python manual_tests/select_text_e2e/test_basic_select.py --test 1
```

## 测试文件说明

### `test_basic_select.py` - 基础功能测试
**测试数量：4 个**

| 测试 | 场景 | 验证要点 |
|-----|------|---------|
| Test 1 | 简单选中文本 | matchCount > 0, selectedText 正确 |
| Test 2 | 选择第N个匹配项 | matchCount ≥ selectIndex, 选中文本正确 |
| Test 3 | 不区分大小写 | 匹配 Hello/HELLO/hello |
| Test 4 | 全字匹配 | 不匹配 test123, mytest |

**运行：**
```bash
uv run python manual_tests/select_text_e2e/test_basic_select.py --test <1-4|all>
```

---

### `test_search_options.py` - 搜索选项测试
**测试数量：5 个**

| 测试 | 场景 | 验证要点 |
|-----|------|---------|
| Test 1 | matchCase: true | 大小写敏感匹配 |
| Test 2 | matchCase: false | 大小写不敏感 |
| Test 3 | matchWholeWord | 完整单词匹配 |
| Test 4 | matchWildcards | 通配符模式（test*） |
| Test 5 | 组合选项 | 多个选项同时启用 |

**运行：**
```bash
uv run python manual_tests/select_text_e2e/test_search_options.py --test <1-5|all>
```

---

### `test_selection_modes.py` - 选择模式测试
**测试数量：4 个**

| 测试 | 场景 | 验证要点 |
|-----|------|---------|
| Test 1 | select 模式 | 文本高亮选中 |
| Test 2 | start 模式 | 光标在文本开头 |
| Test 3 | end 模式 | 光标在文本结尾 |
| Test 4 | 模式切换 | 不同模式间切换 |

**运行：**
```bash
uv run python manual_tests/select_text_e2e/test_selection_modes.py --test <1-4|all>
```

---

### `test_edge_cases.py` - 边界情况测试
**测试数量：5 个**

| 测试 | 场景 | 验证要点 |
|-----|------|---------|
| Test 1 | 未找到匹配 | success=false, matchCount=0 |
| Test 2 | 空搜索文本 | 返回错误或警告 |
| Test 3 | 超出索引范围 | success=false, 显示实际 matchCount |
| Test 4 | 特殊字符 | 正确处理 @#$%, [], {}, () |
| Test 5 | 长文本 | 长文本不截断 |

**运行：**
```bash
uv run python manual_tests/select_text_e2e/test_edge_cases.py --test <1-5|all>
```

## 验证要点

### ✅ 功能验证
- [x] 文本被正确选中（高亮显示）
- [x] 光标位置正确（start/end 模式）
- [x] 选区信息准确（start, end, text）
- [x] 匹配计数正确
- [x] 索引选择正确

### ✅ 数据验证
- [x] **业务层 success 字段检查** - 修复：未找到匹配时 success=false
- [x] `matchCount` 正确
- [x] `selectedIndex` 正确（1-based）
- [x] `selectedText` 内容正确
- [x] `selectionInfo` 结构完整

### ✅ 错误处理验证
- [x] 未找到匹配时 `success=false, matchCount=0`
- [x] 空搜索文本返回错误
- [x] 索引超出范围时操作失败
- [x] 错误消息清晰

## 最近修复

### 2026-01-20 - Add-In 错误响应格式修复
**问题**：未找到文本时，Add-In 返回 `{success: true, data: {success: false}}`，导致 Python 端收到 `error=None`

**修复**（Add-In 端）：
```typescript
// SelectTextHandler.execute() (word-handlers.ts:895-899)
if (!result.success) {
  throw new Error(`Text not found: "${searchText}" (${result.matchCount} matches)`);
}
```

**修复后响应**：
```json
{
  "success": false,
  "error": {
    "code": "3000",
    "message": "Text not found: \"xxx\" (0 matches)"
  }
}
```

**Python 端适配**：
- 更新 `test_edge_cases.py` 验证错误消息格式
- 创建 `test_helpers.py` 统一测试辅助函数
- 所有测试现在返回 `(success, data, error)` 三元组

---

### 2026-01-20 - 验证逻辑增强
1. **修复 `office_workspace.py` success 判断逻辑**
   - 检查业务层的 `success` 字段（而不只是协议层）
   - 当 `matchCount=0` 时正确返回 `success=false`

2. **增强测试验证**
   - 所有测试现在返回 `(success, data, error)` 元组
   - 验证 `matchCount > 0` 确保真正选中
   - 提供详细的错误诊断信息

3. **修复示例**
   ```python
   # 修复前：即使未找到也显示"测试通过"
   success = await select_text(...)
   if success:
       print("✅ 测试通过")  # ❌ 错误

   # 修复后：验证错误消息
   success, data, error = await select_text(...)
   if success:
       print("✅ 测试通过")
   else:
       print(f"❌ 测试失败: {error}")  # ✅ 正确
   ```

## 未实现场景

### 场景 14: 中文字符搜索
**状态：** ❌ 未实现

**测试内容：**
```python
search_text = "中文测试"
success, data = await select_text(workspace, document_uri, search_text)
# 验证中文字符正确编码和匹配
```

**实现建议：**
- 添加到 `test_edge_cases.py` 或新建 `test_i18n.py`
- 验证 UTF-8 编码
- 测试中英文混合搜索

---

### 场景 16: 数字搜索
**状态：** ❌ 未实现

**测试内容：**
```python
search_text = "12345"
success, data = await select_text(workspace, document_uri, search_text)
# 验证数字字符正确匹配
```

**实现建议：**
- 添加到 `test_basic_select.py` 或 `test_edge_cases.py`
- 验证数字不被当作特殊字符
- 测试数字与文本混合（如 "test123"）

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI/pages/42467331) - word:select:text 事件规范
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI/pages/29655312) - 完整协议规范
- [测试清单](../.claude/skills/test-checklist/SKILL.md) - 测试架构说明

## 最后更新

2026-01-20 - 修复验证逻辑，更新 README 结构
