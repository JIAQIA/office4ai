# PPT Delete Element E2E Tests

`ppt:delete:element` 事件的端到端测试。有状态工作流: insert -> get elementId -> delete -> verify。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 单个删除 | empty.pptx | 插入矩形后删除 |
| 2 | 批量删除 | empty.pptx | 插入 3 个矩形后删除 2 个 |
| 3 | 指定幻灯片删除 | multi_element.pptx | 删除 multi_element 的一个元素 |

## 运行

```bash
uv run python manual_tests/ppt/delete_element_e2e/test_delete.py --test 1
uv run python manual_tests/ppt/delete_element_e2e/test_delete.py --test all
uv run python manual_tests/ppt/delete_element_e2e/test_delete.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
