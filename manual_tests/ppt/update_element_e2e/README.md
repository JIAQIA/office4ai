# PPT Update Element E2E Tests

`ppt:update:element` 事件的端到端测试。有状态工作流: insert -> get elementId -> update -> verify。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 移动 | empty.pptx | 插入矩形后移动到 (400, 300) |
| 2 | 缩放 | empty.pptx | 插入矩形后缩放到 400x300 |
| 3 | 旋转 | empty.pptx | 插入矩形后旋转 45 度 |

## 运行

```bash
uv run python manual_tests/ppt/update_element_e2e/test_element_update.py --test 1
uv run python manual_tests/ppt/update_element_e2e/test_element_update.py --test all
uv run python manual_tests/ppt/update_element_e2e/test_element_update.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
