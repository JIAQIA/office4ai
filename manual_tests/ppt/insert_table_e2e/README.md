# PPT Insert Table E2E Tests

`ppt:insert:table` 事件的端到端测试。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 基本表格 | empty.pptx | 插入 3x3 空表格 |
| 2 | 带数据表格 | empty.pptx | 插入预填充数据的 3x2 表格 |
| 3 | 带位置表格 | empty.pptx | 在指定位置插入表格 |

## 运行

```bash
uv run python manual_tests/ppt/insert_table_e2e/test_table_insert.py --test 1
uv run python manual_tests/ppt/insert_table_e2e/test_table_insert.py --test all
uv run python manual_tests/ppt/insert_table_e2e/test_table_insert.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
