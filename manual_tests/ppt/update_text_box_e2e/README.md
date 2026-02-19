# PPT Update TextBox E2E Tests

`ppt:update:textBox` 事件的端到端测试。有状态工作流: insert -> get elementId -> update -> verify。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 修改文本 | empty.pptx | 插入文本框后修改文本内容 |
| 2 | 字体样式 | empty.pptx | 插入文本框后设置 bold/italic |
| 3 | 字号字体 | empty.pptx | 插入文本框后修改 fontSize/fontName |
| 4 | 颜色 | empty.pptx | 插入文本框后修改 color/fillColor |

## 运行

```bash
uv run python manual_tests/ppt/update_text_box_e2e/test_text_box_update.py --test 1
uv run python manual_tests/ppt/update_text_box_e2e/test_text_box_update.py --test all
uv run python manual_tests/ppt/update_text_box_e2e/test_text_box_update.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
