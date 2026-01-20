# Get Selection E2E Tests

`word:get:selection` 事件的手动端到端测试。

## 测试场景

### 测试 1: 获取正常选区（有高亮文本）
获取文档中的高亮文本选区信息。

**前置操作:**
1. 在 Word 文档中输入一些文本
2. 用鼠标选中部分文本（创建高亮选区）

**预期结果:**
- `isEmpty` 为 `false`
- `type` 为 `Normal`
- `start` 和 `end` 有值（字符偏移）
- `text` 包含选中的文本内容

### 测试 2: 获取光标位置（无高亮文本）
获取光标插入点位置信息。

**前置操作:**
1. 在 Word 文档中点击，确保光标在某个位置
2. 不要选中任何文本

**预期结果:**
- `isEmpty` 为 `true`
- `type` 为 `InsertionPoint`
- `start` 和 `end` 值相同（光标位置）
- `text` 为 `null` 或不存在

### 测试 3: 获取无选区状态
获取无选区时的状态信息。

**前置操作:**
1. 确保文档没有任何选中内容
2. 光标不在编辑区域（可能在文档外）

**预期结果:**
- `isEmpty` 为 `true`
- `type` 为 `NoSelection`
- `start` 和 `end` 为 `null`
- `text` 为 `null`

### 测试 4: 验证轻量级特性
验证 word:get:selection 仅返回位置信息，不返回完整内容结构。

**对比验证:**
- `word:get:selection` 仅返回：`isEmpty`, `type`, `start`, `end`, `text`
- `word:get:selectedContent` 返回：完整内容结构（段落、表格、图片等）

**预期结果:**
- 响应时间明显快于 `word:get:selectedContent`
- 数据量小，仅包含必要的位置信息

## 运行方式

### 前置条件

1. 启动 Workspace 服务器
2. 在 Word 中加载 office-editor4ai Add-In
3. 确保文档已打开并连接

### 运行单个测试

```bash
# 测试 1: 获取正常选区（有高亮文本）
uv run python manual_tests/get_selection_e2e/test_selection.py --test 1

# 测试 2: 获取光标位置（无高亮文本）
uv run python manual_tests/get_selection_e2e/test_selection.py --test 2

# 测试 3: 获取无选区状态
uv run python manual_tests/get_selection_e2e/test_selection.py --test 3

# 测试 4: 验证轻量级特性
uv run python manual_tests/get_selection_e2e/test_selection.py --test 4
```

### 运行所有测试

```bash
uv run python manual_tests/get_selection_e2e/test_selection.py --test all
```

## 验证要点

### 功能验证
- [ ] 能成功获取选区状态
- [ ] 选区类型正确（Normal/InsertionPoint/NoSelection）
- [ ] 位置信息准确（start/end 值）
- [ ] 文本内容正确（仅在有选区时）

### 数据验证
- [ ] `isEmpty` 字段与实际选区状态一致
- [ ] `type` 字段值在枚举范围内
- [ ] `start` ≤ `end`（当两者都有值时）
- [ ] `text` 长度 = `end` - `start`（Normal 类型）

### 性能验证
- [ ] 响应时间 < 1 秒（轻量级查询）
- [ ] 数据量小，无明显延迟

### 用户体验
- [ ] 错误提示清晰
- [ ] 状态描述易于理解
- [ ] 光标位置显示准确

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/36569100) - word:get:selection 事件规范
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范
- [Word 事件索引](https://turingfocus.atlassian.net/wiki/pages/29851649) - 完整事件列表

## 最后更新

2026-01-14
