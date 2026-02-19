# PPT Insert Image E2E Tests

`ppt:insert:image` 事件的端到端测试。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 基本图片插入 | empty.pptx | 在空白幻灯片插入最小 PNG 图片 |
| 2 | 带位置图片插入 | empty.pptx | 在指定位置 (200, 150) 插入图片 |
| 3 | 指定幻灯片插入 | simple.pptx | 在 slideIndex=1 上插入图片 |

## 运行

```bash
uv run python manual_tests/ppt/insert_image_e2e/test_image_insert.py --test 1
uv run python manual_tests/ppt/insert_image_e2e/test_image_insert.py --test all
uv run python manual_tests/ppt/insert_image_e2e/test_image_insert.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
