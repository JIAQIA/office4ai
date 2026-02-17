# PPT Get Screenshot E2E Tests

`ppt:get:slideScreenshot` 事件的端到端测试。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | PNG 截图 | colored_slide.pptx | 以默认 PNG 格式获取彩色幻灯片截图 |
| 2 | JPEG 截图 | colored_slide.pptx | 以 JPEG 格式获取截图 |
| 3 | Base64 格式验证 | colored_slide.pptx | 验证截图返回数据是合法的 Base64 编码 |

## 运行

```bash
uv run python manual_tests/ppt_get_screenshot_e2e/test_screenshot.py --test 1
uv run python manual_tests/ppt_get_screenshot_e2e/test_screenshot.py --test all
uv run python manual_tests/ppt_get_screenshot_e2e/test_screenshot.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
