# Get Styles E2E Tests

`word:get:styles` 事件的手动端到端测试。

## 测试场景

### 测试 1: 获取所有正在使用的样式（默认参数）
获取文档中所有正在使用的样式，包括内置和自定义样式。

**预期结果:**
- 返回所有正在使用的样式
- 包括段落、字符、表格、列表类型
- 按类型分组显示

### 测试 2: 仅获取内置样式
只获取 Word 内置的样式。

**参数:**
```python
{
    "includeBuiltIn": True,
    "includeCustom": False,
    "includeUnused": False,
    "detailedInfo": False
}
```

**预期结果:**
- 仅返回内置样式（如"标题一"、"正文"等）
- 不包括自定义样式

### 测试 3: 仅获取自定义样式
只获取用户自定义的样式。

**参数:**
```python
{
    "includeBuiltIn": False,
    "includeCustom": True,
    "includeUnused": True,
    "detailedInfo": False
}
```

**预期结果:**
- 仅返回自定义样式
- 包括未使用的自定义样式

### 测试 4: 获取包含详细信息的样式
获取样式及其详细描述信息。

**参数:**
```python
{
    "includeBuiltIn": True,
    "includeCustom": True,
    "includeUnused": False,
    "detailedInfo": True
}
```

**预期结果:**
- 返回样式时包含描述字段
- 显示每个样式的详细信息

### 测试 5: 获取所有样式（包括未使用的）
获取文档中的所有样式，包括未使用的。

**参数:**
```python
{
    "includeBuiltIn": True,
    "includeCustom": True,
    "includeUnused": True,
    "detailedInfo": False
}
```

**预期结果:**
- 返回所有样式，包括未使用的
- 数量可能比测试 1 多

## 运行方式

### 前置条件

1. 启动 Workspace 服务器
2. 在 Word 中加载 office-editor4ai Add-In
3. 确保文档已打开并连接

### 运行单个测试

```bash
# 测试 1: 获取所有正在使用的样式
uv run python manual_tests/get_styles_e2e/test_styles.py --test 1

# 测试 2: 仅获取内置样式
uv run python manual_tests/get_styles_e2e/test_styles.py --test 2

# 测试 3: 仅获取自定义样式
uv run python manual_tests/get_styles_e2e/test_styles.py --test 3

# 测试 4: 获取包含详细信息的样式
uv run python manual_tests/get_styles_e2e/test_styles.py --test 4

# 测试 5: 获取所有样式（包括未使用的）
uv run python manual_tests/get_styles_e2e/test_styles.py --test 5
```

### 运行所有测试

```bash
uv run python manual_tests/get_styles_e2e/test_styles.py --test all
```

## 验证要点

### 功能验证
- [ ] 能成功获取样式列表
- [ ] 样式类型正确（Paragraph/Character/Table/List）
- [ ] 内置/自定义标识正确
- [ ] 使用状态标识正确
- [ ] 过滤选项生效

### 数据验证
- [ ] 样式名称为本地化名称（如"标题一"而非"Heading 1"）
- [ ] 按类型分组统计正确
- [ ] 内置和自定义样式数量合理
- [ ] 详细信息仅在请求时返回

### 用户体验
- [ ] 响应时间合理（< 5 秒）
- [ ] 错误提示清晰
- [ ] 显示格式易于阅读

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/32702465) - word:get:styles 事件规范
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范
- [Word 事件索引](https://turingfocus.atlassian.net/wiki/pages/29851649) - 完整事件列表
