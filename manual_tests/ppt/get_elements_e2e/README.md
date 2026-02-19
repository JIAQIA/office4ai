# PPT Get Elements E2E Tests

`ppt:get:currentSlideElements` 和 `ppt:get:slideElements` 事件的端到端测试。

## 测试文件

### test_current_slide.py — 当前幻灯片元素

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 空幻灯片元素 | empty.pptx | 获取空白幻灯片的元素，验证返回格式正确 |
| 2 | 多元素幻灯片 | multi_element.pptx | 获取含文本框+表格+形状的幻灯片元素 |
| 3 | 元素属性验证 | multi_element.pptx | 验证每个元素都包含 id 和 type 字段 |

### test_specific_slide.py — 指定幻灯片元素

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 按索引获取元素 | multi_element.pptx | 获取 slideIndex=0 的所有元素 |
| 2 | 过滤文本元素 | multi_element.pptx | includeText=true, includeShapes=false |
| 3 | 过滤形状元素 | multi_element.pptx | includeShapes=true, includeText=false |

## 运行

```bash
# 当前幻灯片
uv run python manual_tests/ppt/get_elements_e2e/test_current_slide.py --test all

# 指定幻灯片
uv run python manual_tests/ppt/get_elements_e2e/test_specific_slide.py --test all
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
