# Get Styles E2E Tests

`word:get:styles` 事件的端到端测试（自动化版本）。

## 测试场景

| 测试编号 | 测试名称 | 选项 | 描述 |
|---------|---------|------|------|
| 1 | 获取所有正在使用的样式 | 默认 | 获取文档中所有正在使用的样式 |
| 2 | 仅获取内置样式 | `includeBuiltIn=true, includeCustom=false` | 只获取 Word 内置样式 |
| 3 | 仅获取自定义样式 | `includeBuiltIn=false, includeCustom=true` | 只获取用户自定义样式 |
| 4 | 获取包含详细信息的样式 | `detailedInfo=true` | 获取样式及其详细描述信息 |
| 5 | 获取所有样式（包括未使用的） | `includeUnused=true` | 获取所有样式，包括未使用的 |

## 运行方式

### 前置条件

1. 在 Word 中加载 office-editor4ai Add-In
2. 测试会自动启动 Workspace 服务器
3. 测试会自动打开文档并等待 Add-In 连接

### 运行单个测试

```bash
# 测试 1: 获取所有正在使用的样式（默认参数）
uv run python manual_tests/word/get_styles_e2e/test_styles.py --test 1

# 测试 2: 仅获取内置样式
uv run python manual_tests/word/get_styles_e2e/test_styles.py --test 2

# 测试 3: 仅获取自定义样式
uv run python manual_tests/word/get_styles_e2e/test_styles.py --test 3

# 测试 4: 获取包含详细信息的样式
uv run python manual_tests/word/get_styles_e2e/test_styles.py --test 4

# 测试 5: 获取所有样式（包括未使用的）
uv run python manual_tests/word/get_styles_e2e/test_styles.py --test 5
```

### 运行所有测试

```bash
uv run python manual_tests/word/get_styles_e2e/test_styles.py --test all
```

### 其他选项

```bash
# 列出所有测试用例
uv run python manual_tests/word/get_styles_e2e/test_styles.py --list

# 手动打开文档模式（不自动打开）
uv run python manual_tests/word/get_styles_e2e/test_styles.py --test 1 --no-auto-open

# 失败时也清理测试文件
uv run python manual_tests/word/get_styles_e2e/test_styles.py --test 1 --always-cleanup
```

## 测试流程

1. 自动复制夹具文件到临时目录
2. 自动打开文档（或手动打开）
3. 等待 Add-In 连接
4. 执行 `word:get:styles` 事件
5. 验证返回数据结构和过滤逻辑
6. 成功后自动清理，失败则保留供调试

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

### 性能验证
- [ ] 响应时间合理（< 5 秒）

## 夹具文件

测试使用 `fixtures/get_styles_e2e/` 目录下的文件：
- `simple.docx` - 简单文本文档
- `complex.docx` - 复杂文档（包含表格、列表等）

夹具文件会在首次运行时自动创建。

## 选项参数说明

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `includeBuiltIn` | boolean | true | 是否包含内置样式 |
| `includeCustom` | boolean | true | 是否包含自定义样式 |
| `includeUnused` | boolean | false | 是否包含未使用的样式 |
| `detailedInfo` | boolean | false | 是否返回详细描述信息 |

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/32702465) - word:get:styles 事件规范
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范
- [Word 事件索引](https://turingfocus.atlassian.net/wiki/pages/29851649) - 完整事件列表

## 最后更新

2026-02-04
