# Comment E2E Tests

`word:insert:comment` / `word:get:comments` / `word:reply:comment` / `word:resolve:comment` / `word:delete:comment` 事件的端到端测试（自动化版本）。

## 测试文件

### test_comment_crud.py — CRUD 工作流

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 插入+获取批注 | insert:comment → get:comments, 验证批注内容 |
| 2 | 插入+回复+获取批注 | insert:comment → reply:comment → get:comments, 验证回复 |
| 3 | 插入+解决+获取批注 | insert:comment → resolve:comment → get:comments, 验证 resolved=True |
| 4 | 插入+删除+获取批注 | insert:comment → delete:comment → get:comments, 验证批注已删除 |

### test_comment_target.py — 批注定位

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 默认选区批注 | target=None (使用当前选区) |
| 2 | 搜索文本批注 | target={type:"searchText", searchText:"测试"} |
| 3 | 搜索不存在文本 | target={type:"searchText", searchText:"不存在的文本"}, 预期失败 |

### test_comment_options.py — 查询选项

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 排除已解决批注 | includeResolved=false 不返回已解决批注 |
| 2 | 包含关联文本 | includeAssociatedText=true, 验证 associatedText 字段 |
| 3 | 详细元数据 | detailedMetadata=true, 验证 authorName/creationDate |

## 运行方式

### 前置条件

1. 在 Word 中加载 office-editor4ai Add-In
2. 测试会自动启动 Workspace 服务器
3. 测试会自动打开文档并等待 Add-In 连接

### CRUD 工作流测试

```bash
# 运行单个测试
uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test 1
uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test 2
uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test 3
uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test 4

# 运行全部
uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test all

# 列出测试
uv run python manual_tests/word/comment_e2e/test_comment_crud.py --list
```

### 批注定位测试

```bash
uv run python manual_tests/word/comment_e2e/test_comment_target.py --test 1
uv run python manual_tests/word/comment_e2e/test_comment_target.py --test 2
uv run python manual_tests/word/comment_e2e/test_comment_target.py --test 3

uv run python manual_tests/word/comment_e2e/test_comment_target.py --test all
```

### 查询选项测试

```bash
uv run python manual_tests/word/comment_e2e/test_comment_options.py --test 1
uv run python manual_tests/word/comment_e2e/test_comment_options.py --test 2
uv run python manual_tests/word/comment_e2e/test_comment_options.py --test 3

uv run python manual_tests/word/comment_e2e/test_comment_options.py --test all
```

### 其他选项

```bash
# 手动打开文档模式（不自动打开）
uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test 1 --no-auto-open

# 失败时也清理测试文件
uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test 1 --always-cleanup
```

## 测试流程

1. 自动复制夹具文件到临时目录
2. 自动打开文档（或手动打开）
3. 等待 Add-In 连接
4. 执行链式批注操作（有状态工作流）
5. 验证返回数据和文档批注状态
6. 成功后自动清理，失败则保留供调试

## 验证要点

### CRUD 验证
- [ ] insert:comment 成功返回 commentId
- [ ] get:comments 列表中包含新插入的批注
- [ ] reply:comment 成功后 replies 中包含回复内容
- [ ] resolve:comment 成功后 resolved=True
- [ ] delete:comment 成功后批注从列表中消失

### 定位验证
- [ ] 默认选区批注附加到当前光标位置
- [ ] searchText 定位批注关联到匹配文本
- [ ] 搜索不存在文本时失败或降级处理

### 选项验证
- [ ] includeResolved=false 正确过滤已解决批注
- [ ] includeAssociatedText=true 返回 associatedText 字段
- [ ] detailedMetadata=true 返回 authorName/creationDate

## 夹具文件

测试使用 `fixtures/comment_e2e/` 目录下的文件：
- `simple.docx` - 简单文本文档

夹具文件会在首次运行时自动创建。

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/32702465) - Word 批注事件规范
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范

## 最后更新

2026-02-19
