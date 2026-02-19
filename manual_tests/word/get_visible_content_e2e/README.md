# Get Visible Content E2E Tests

`word:get:visibleContent` 事件的手动端到端测试。

## 测试场景

### 基础测试 (`test_basic_get.py`)

测试基础的可见内容获取功能。

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 获取可见文本 | 获取当前可见区域的文本内容 |
| 2 | 获取空文档 | 获取空文档的可见内容 |
| 3 | 获取格式化文本 | 获取包含格式化文本的可见内容 |
| 4 | 获取混合元素 | 获取包含文本、图片、表格的可见内容 |

### 选项测试 (`test_options_get.py`)

测试各种选项参数组合。

| 测试编号 | 测试名称 | options 参数 | 描述 |
|---------|---------|-------------|------|
| 1 | 文本和图片 | `{includeText: true, includeImages: true, includeTables: false}` | 获取文本和图片 |
| 2 | 文本和表格 | `{includeText: true, includeImages: false, includeTables: true}` | 获取文本和表格 |
| 3 | 文本长度限制 | `{includeText: true, maxTextLength: 100}` | 限制单个元素（段落/单元格）的文本长度 |
| 4 | 仅文本 | `{includeText: true, includeImages: false, includeTables: false}` | 仅获取文本 |
| 5 | 详细元数据 | `{includeText: true, detailedMetadata: true}` | 获取详细元数据 |

### 边界测试 (`test_edge_cases.py`)

测试边界情况和异常场景。

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 超长文档 | 获取超长文档的可见内容 |
| 2 | 特殊字符 | 获取包含特殊字符的内容 |
| 3 | 嵌入对象 | 获取包含嵌入对象的内容 |
| 4 | 连续请求 | 多次连续获取可见内容 |

## 运行方式

### 前置条件

1. 启动 Workspace 服务器（测试会自动启动）
2. 在 Word 中加载 office-editor4ai Add-In
3. 确保文档已打开并连接

### 运行单个测试

```bash
# 基础测试
uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test 1
uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test 2
uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test 3
uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test 4

# 选项测试
uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test 1
uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test 2
uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test 3
uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test 4
uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test 5

# 边界测试
uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test 1
uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test 2
uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test 3
uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test 4
```

### 运行所有测试

```bash
# 基础测试（4个）
uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test all

# 选项测试（5个）
uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test all

# 边界测试（4个）
uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test all
```

## 验证要点

### 功能验证
- [ ] 能成功获取可见内容
- [ ] 文本内容正确返回
- [ ] 元数据正确（isEmpty, characterCount）
- [ ] 元素列表正确返回
- [ ] 选项参数生效

### 数据验证
- [ ] 文本长度符合预期
- [ ] 元素类型正确（text/image/table/other）
- [ ] maxTextLength 限制单个元素长度生效（段落/单元格/控件）
- [ ] includeImages/includeTables 过滤生效

### 用户体验
- [ ] 响应时间合理（< 5 秒）
- [ ] 错误提示清晰
- [ ] 显示格式易于阅读

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/30736386) - word:get:visibleContent 事件规范
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范
- [Word 事件索引](https://turingfocus.atlassian.net/wiki/pages/29851649) - 完整事件列表

## 最后更新

2026-01-12
