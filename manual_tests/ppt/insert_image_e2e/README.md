# PPT Insert Image E2E Tests

`ppt:insert:image` 事件的端到端测试。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 基本图片插入 | empty.pptx | 在空白幻灯片插入最小 PNG 图片 |
| 2 | 带位置图片插入 | empty.pptx | 在指定位置 (200, 150) 插入图片 |
| 3 | 跨幻灯片插入 | simple.pptx | slideIndex=1，验证目标幻灯片有新元素 |
| 4 | slideIndex 越界 | simple.pptx | slideIndex=999，预期返回失败 |
| 5 | 视图恢复验证 | simple.pptx | 跨页插入后 currentSlideIndex 应恢复 |

### 跨幻灯片测试说明 (Tests 3-5)

PowerPoint JS API 没有 `slide.shapes.addImage()` 方法，图片插入只能通过 `setSelectedDataAsync`，
该 API 固定作用于当前选中幻灯片。Add-In 采用 **goto-then-insert** 方案：

1. 导航到 `slideIndex` 目标幻灯片
2. `setSelectedDataAsync` 插入图片
3. 导航回原始幻灯片（finally 块）

因此 test 5 的视图恢复验证覆盖了 Add-In 的 goto-back 逻辑是否正确。

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
