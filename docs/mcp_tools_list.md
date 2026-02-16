# MCP Tools 列表

> **与 TypeScript Socket.IO API 一对一映射**
> **Python Pydantic 模型与 TS 类型严格同步**

---

## 命名规范

采用 `<platform>_<action>[:<resource>]` 命名规范，对应 Socket.IO 事件：
- Socket.IO: `word:insert:text` → MCP Tool: `word_insert_text`
- Socket.IO: `ppt:get:currentSlideElements` → MCP Tool: `ppt_get_current_slide_elements`

---

## Word Tools (13 个)

### 内容获取类 (4 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `word_get_selected_content` | `word:get:selectedContent` | 获取选中内容 |
| `word_get_visible_content` | `word:get:visibleContent` | 获取可见内容 |
| `word_get_document_structure` | `word:get:documentStructure` | 获取文档结构 |
| `word_get_document_stats` | `word:get:documentStats` | 获取统计信息 |

### 文本操作类 (4 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `word_insert_text` | `word:insert:text` | 插入文本 |
| `word_replace_selection` | `word:replace:selection` | 替换选中内容 |
| `word_replace_text` | `word:replace:text` | 查找替换文本 |
| `word_append_text` | `word:append:text` | 追加文本 |

### 多模态操作类 (3 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `word_insert_image` | `word:insert:image` | 插入图片 |
| `word_insert_table` | `word:insert:table` | 插入表格 |
| `word_insert_equation` | `word:insert:equation` | 插入公式 |

### 高级功能类 (2 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `word_insert_toc` | `word:insert:toc` | 插入目录 |
| `word_export_content` | `word:export:content` | 导出内容 |

---

## PowerPoint Tools (21 个)

### 内容获取类 (5 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `ppt_get_current_slide_elements` | `ppt:get:currentSlideElements` | 获取当前页元素 |
| `ppt_get_slide_elements` | `ppt:get:slideElements` | 获取指定页元素 |
| `ppt_get_slide_screenshot` | `ppt:get:slideScreenshot` | 获取幻灯片截图 |
| `ppt_get_slide_info` | `ppt:get:slideInfo` | 获取演示文稿/幻灯片基本信息 |
| `ppt_get_slide_layouts` | `ppt:get:slideLayouts` | 获取可用幻灯片版式列表 |

### 内容插入类 (4 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `ppt_insert_text` | `ppt:insert:text` | 插入文本框 |
| `ppt_insert_image` | `ppt:insert:image` | 插入图片 |
| `ppt_insert_table` | `ppt:insert:table` | 插入表格 |
| `ppt_insert_shape` | `ppt:insert:shape` | 插入形状 |

### 更新操作类 (6 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `ppt_update_text_box` | `ppt:update:textBox` | 更新文本框内容/样式 |
| `ppt_update_image` | `ppt:update:image` | 替换图片内容 |
| `ppt_update_table_cell` | `ppt:update:tableCell` | 更新表格单元格 |
| `ppt_update_table_row_column` | `ppt:update:tableRowColumn` | 按行/列批量更新表格 |
| `ppt_update_table_format` | `ppt:update:tableFormat` | 更新表格样式格式 |
| `ppt_update_element` | `ppt:update:element` | 更新元素位置/尺寸/旋转 |

### 删除与布局类 (2 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `ppt_delete_element` | `ppt:delete:element` | 删除元素 |
| `ppt_reorder_element` | `ppt:reorder:element` | 调整元素层叠顺序 |

### 幻灯片管理类 (4 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `ppt_add_slide` | `ppt:add:slide` | 添加幻灯片 |
| `ppt_delete_slide` | `ppt:delete:slide` | 删除幻灯片 |
| `ppt_move_slide` | `ppt:move:slide` | 移动幻灯片 |
| `ppt_goto_slide` | `ppt:goto:slide` | 跳转到幻灯片 |

---

## Excel Tools (4 个)

| MCP 工具名称 | Socket.IO 事件 | 功能描述 |
|-------------|---------------|----------|
| `excel_get_selected_range` | `excel:get:selectedRange` | 获取选中范围 |
| `excel_set_cell_value` | `excel:set:cellValue` | 设置单元格值 |
| `excel_insert_table` | `excel:insert:table` | 插入表格 |
| `excel_get_used_range` | `excel:get:usedRange` | 获取已使用范围 |

---

## 汇总

| 类别 | 工具数量 | 对应 TypeScript 文件 |
|------|---------|---------------------|
| Word | 13 | `word-editor4ai/src/word-tools/` |
| PowerPoint | 21 | `ppt-editor4ai/` (待确认路径) |
| Excel | 4 | `excel-editor4ai/` (待确认路径) |
| **总计** | **38** | |

---

## 实现映射

每个 MCP Tool 的实现对应一个 Python 文件：

```
office4ai/a2c_smcp/tools/
├── word/
│   ├── get_selected_content.py    # word_get_selected_content
│   ├── insert_text.py              # word_insert_text
│   ├── replace_selection.py        # word_replace_selection
│   └── ... (其余 10 个)
├── ppt/
│   ├── get_current_slide_elements.py  # ppt_get_current_slide_elements
│   ├── insert_text.py                  # ppt_insert_text
│   └── ... (其余 19 个)
└── excel/
    ├── get_selected_range.py       # excel_get_selected_range
    ├── set_cell_value.py           # excel_set_cell_value
    └── ... (其余 2 个)
```
