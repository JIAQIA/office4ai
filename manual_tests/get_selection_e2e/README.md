# Get Selection E2E Tests

`word:get:selection` 事件的端到端测试（自动化版本）。

## 测试场景

| 测试编号 | 测试名称 | 用户操作 | 描述 |
|---------|---------|---------|------|
| 1 | 正常选区（有高亮文本） | 选中一段文本 | 验证 `isEmpty=false`, `type=Normal`, 有 `text` |
| 2 | 光标位置（无高亮文本） | 点击放置光标 | 验证 `isEmpty=true`, `type=InsertionPoint`, `start=end` |
| 3 | 无选区状态 | 点击文档外部 | 验证 `isEmpty=true`, `type=NoSelection` |
| 4 | 性能对比 | 选中一大段文本 | 对比 `get:selection` vs `get:selectedContent` 性能 |

## 运行方式

### 前置条件

1. 在 Word 中加载 office-editor4ai Add-In
2. 测试会自动启动 Workspace 服务器
3. 测试会自动打开文档并等待 Add-In 连接

### 运行单个测试

```bash
# 测试 1: 正常选区（有高亮文本）
uv run python manual_tests/get_selection_e2e/test_selection.py --test 1

# 测试 2: 光标位置（无高亮文本）
uv run python manual_tests/get_selection_e2e/test_selection.py --test 2

# 测试 3: 无选区状态
uv run python manual_tests/get_selection_e2e/test_selection.py --test 3

# 测试 4: 性能对比
uv run python manual_tests/get_selection_e2e/test_selection.py --test 4
```

### 运行所有测试

```bash
uv run python manual_tests/get_selection_e2e/test_selection.py --test all
```

### 其他选项

```bash
# 列出所有测试用例
uv run python manual_tests/get_selection_e2e/test_selection.py --list

# 手动打开文档模式（不自动打开）
uv run python manual_tests/get_selection_e2e/test_selection.py --test 1 --no-auto-open

# 失败时也清理测试文件
uv run python manual_tests/get_selection_e2e/test_selection.py --test 1 --always-cleanup
```

## 测试流程

1. 自动复制夹具文件到临时目录
2. 自动打开文档（或手动打开）
3. 等待 Add-In 连接
4. **提示用户执行选区操作**（如选中文本）
5. 按 Enter 继续执行测试
6. 验证协议返回数据
7. 成功后自动清理，失败则保留供调试

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
- [ ] `text` 长度 ≈ `end` - `start`（Normal 类型，可能因编码有差异）

### 性能验证
- [ ] `get:selection` 响应时间 < 1 秒
- [ ] `get:selection` 比 `get:selectedContent` 更快

## 平台兼容性

`type` 字段的获取方式因平台而异：

| 平台 | `type` 获取方式 |
|------|----------------|
| Windows Desktop | 直接使用 `Selection.type` API (WordApiDesktop 1.4) |
| Word Online / Mac | 基于 `isEmpty` 推断：`true` → `InsertionPoint`，`false` → `Normal` |

**注意**：`NoSelection` 类型仅在 Windows Desktop 且光标完全不在文档内时返回。

## 夹具文件

测试使用 `fixtures/get_selection_e2e/` 目录下的文件：
- `simple.docx` - 简单文本文档
- `large.docx` - 大型文档（性能测试）

夹具文件会在首次运行时自动创建。

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/36569100) - word:get:selection 事件规范
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范
- [Word 事件索引](https://turingfocus.atlassian.net/wiki/pages/29851649) - 完整事件列表

## 最后更新

2026-02-04
