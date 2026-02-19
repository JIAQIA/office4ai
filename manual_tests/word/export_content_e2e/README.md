# Export Content E2E Tests

`word:export:content` 事件的端到端测试（自动化版本）。

## 测试文件

### test_basic_export.py — 基本导出

| 测试编号 | 测试名称 | 格式 | 描述 |
|---------|---------|------|------|
| 1 | 纯文本导出 | text | 验证 content 非空且包含已知文本 |
| 2 | HTML 导出 | html | 验证 content 包含 HTML 标签 |
| 3 | Markdown 导出 | markdown | 验证 content 非空 |
| 4 | 空文档导出 | text | 验证 success=True, content 为空或极短 |

### test_export_options.py — 导出选项

| 测试编号 | 测试名称 | 格式 | 描述 |
|---------|---------|------|------|
| 1 | 包含表格导出 (HTML) | html | 复杂文档导出，验证 content 包含表格标记 |
| 2 | 包含表格选项 (Markdown) | markdown | includeTables=true, 验证表格内容 |
| 3 | 大文档性能 | text | 大文档导出，验证执行时间和 content 长度 > 1000 |

## 运行方式

### 前置条件

1. 在 Word 中加载 office-editor4ai Add-In
2. 测试会自动启动 Workspace 服务器
3. 测试会自动打开文档并等待 Add-In 连接

### 基本导出测试

```bash
# 运行单个测试
uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test 1
uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test 2
uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test 3
uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test 4

# 运行全部
uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test all

# 列出测试
uv run python manual_tests/word/export_content_e2e/test_basic_export.py --list
```

### 导出选项测试

```bash
uv run python manual_tests/word/export_content_e2e/test_export_options.py --test 1
uv run python manual_tests/word/export_content_e2e/test_export_options.py --test 2
uv run python manual_tests/word/export_content_e2e/test_export_options.py --test 3

uv run python manual_tests/word/export_content_e2e/test_export_options.py --test all
```

### 其他选项

```bash
# 手动打开文档模式（不自动打开）
uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test 1 --no-auto-open

# 失败时也清理测试文件
uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test 1 --always-cleanup
```

## 测试流程

1. 自动复制夹具文件到临时目录
2. 自动打开文档（或手动打开）
3. 等待 Add-In 连接
4. 执行 `word:export:content` 事件
5. 验证返回的 content 格式和内容
6. 成功后自动清理，失败则保留供调试

## 验证要点

### 格式验证
- [ ] text 格式返回纯文本，包含文档中的已知文字
- [ ] html 格式返回包含 HTML 标签的内容
- [ ] markdown 格式返回非空内容

### 选项验证
- [ ] 复杂文档的 HTML 导出包含 `<table>` 等表格标记
- [ ] includeTables=true 的 Markdown 导出包含 `|` 分隔符
- [ ] 空文档导出 content 为空或极短

### 性能验证
- [ ] 大文档导出响应时间合理（< 10 秒）
- [ ] 大文档 content 长度 > 1000 字符

## 导出参数说明

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `format` | string | `"text"` | 导出格式: `text`, `html`, `markdown` |
| `options.includeTables` | boolean | true | 是否包含表格内容 |

## 夹具文件

测试使用 `fixtures/export_content_e2e/` 目录下的文件：
- `simple.docx` - 简单文本文档
- `complex.docx` - 复杂文档（包含表格、列表等）
- `empty.docx` - 空文档
- `large.docx` - 大文档

夹具文件会在首次运行时自动创建。

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/32702465) - word:export:content 事件规范
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范

## 最后更新

2026-02-19
