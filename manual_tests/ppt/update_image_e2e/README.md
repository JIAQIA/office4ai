# PPT Update Image E2E Tests

`ppt:update:image` 事件的端到端测试。有状态工作流: insert -> get elementId -> update -> verify。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 替换图片 | empty.pptx | 插入红色 PNG 后替换为蓝色 PNG |
| 2 | keepDimensions | empty.pptx | 替换图片时保持原始尺寸 |

## 运行

```bash
uv run python manual_tests/ppt/update_image_e2e/test_image_update.py --test 1
uv run python manual_tests/ppt/update_image_e2e/test_image_update.py --test all
uv run python manual_tests/ppt/update_image_e2e/test_image_update.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
