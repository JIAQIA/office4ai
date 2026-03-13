# PPT Reorder Element E2E Tests

`ppt:reorder:element` 事件的端到端测试。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | bringToFront | empty.pptx | 将第一个形状移到最前 |
| 2 | sendToBack | empty.pptx | 将第一个形状移到最后 |
| 3 | bringForward | empty.pptx | 将第一个形状前移一层 |
| 4 | sendBackward | empty.pptx | 将第一个形状后移一层 |

## 运行

```bash
uv run python manual_tests/ppt/reorder_element_e2e/test_reorder.py --test 1
uv run python manual_tests/ppt/reorder_element_e2e/test_reorder.py --test all
uv run python manual_tests/ppt/reorder_element_e2e/test_reorder.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
