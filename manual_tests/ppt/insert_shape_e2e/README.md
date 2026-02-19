# PPT Insert Shape E2E Tests

`ppt:insert:shape` 事件的端到端测试。

## 测试文件

### test_shape_insert.py — 基本形状插入

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 插入矩形 | empty.pptx | 插入 Rectangle 形状 |
| 2 | 插入圆形 | empty.pptx | 插入 Circle 形状 |
| 3 | 带文本形状 | empty.pptx | 插入带 text='Hello Shape' 的矩形 |
| 4 | 带样式形状 | empty.pptx | 插入带红色填充和蓝色边框的矩形 |

### test_shape_types.py — 批量形状类型

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 批量基本形状 | empty.pptx | 连续插入 Rectangle, Circle, Oval, Triangle |
| 2 | 箭头线条 | empty.pptx | 插入 Arrow 和 Line |
| 3 | 星形六边形 | empty.pptx | 插入 Star, Hexagon, Pentagon |

## 运行

```bash
# 基本形状
uv run python manual_tests/ppt/insert_shape_e2e/test_shape_insert.py --test all

# 形状类型
uv run python manual_tests/ppt/insert_shape_e2e/test_shape_types.py --test all
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
