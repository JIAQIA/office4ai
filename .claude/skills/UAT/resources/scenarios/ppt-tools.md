# UAT 场景：PPT 工具注册验收 (ppt-tools)

## 测试目标

验证 21 个 PPT MCP 工具在 MCP Inspector 中正确注册，参数 schema 符合预期。

## 验证清单

### 内容读取工具（5 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-01 | `ppt_get_current_slide_elements` | 当前幻灯片元素 | `document_uri` | - |
| P-02 | `ppt_get_slide_elements` | 指定幻灯片元素 | `document_uri`, `slide_index` | - |
| P-03 | `ppt_get_slide_screenshot` | 幻灯片截图 | `document_uri` | `slide_index` |
| P-04 | `ppt_get_slide_info` | 幻灯片信息 | `document_uri` | - |
| P-05 | `ppt_get_slide_layouts` | 版式模板 | `document_uri` | - |

### 元素插入工具（4 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-06 | `ppt_insert_text` | 插入文本框 | `document_uri`, `text` | `left`, `top`, `width`, `height`, `slide_index` |
| P-07 | `ppt_insert_image` | 插入图片 | `document_uri`, `image_source` | `left`, `top`, `width`, `height`, `slide_index` |
| P-08 | `ppt_insert_table` | 插入表格 | `document_uri`, `rows`, `columns` | `left`, `top`, `slide_index` |
| P-09 | `ppt_insert_shape` | 插入形状 | `document_uri`, `shape_type` | `left`, `top`, `width`, `height`, `slide_index` |

### 元素更新工具（6 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-10 | `ppt_update_text_box` | 更新文本框 | `document_uri`, `element_id` | `text`, `style` |
| P-11 | `ppt_update_image` | 更新图片 | `document_uri`, `element_id`, `image_source` | - |
| P-12 | `ppt_update_table_cell` | 更新单元格 | `document_uri`, `element_id`, `row`, `column` | `text`, `style` |
| P-13 | `ppt_update_table_row_column` | 批量更新行列 | `document_uri`, `element_id` | `updates` |
| P-14 | `ppt_update_table_format` | 更新表格格式 | `document_uri`, `element_id` | `format` |
| P-15 | `ppt_update_element` | 更新元素位置/大小 | `document_uri`, `element_id` | `left`, `top`, `width`, `height`, `rotation` |

### 删除与排序工具（2 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-16 | `ppt_delete_element` | 删除元素 | `document_uri`, `element_id` | - |
| P-17 | `ppt_reorder_element` | 调整层级 | `document_uri`, `element_id` | `position` |

### 幻灯片管理工具（4 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-18 | `ppt_add_slide` | 新增幻灯片 | `document_uri` | `layout`, `index` |
| P-19 | `ppt_delete_slide` | 删除幻灯片 | `document_uri`, `slide_index` | - |
| P-20 | `ppt_move_slide` | 移动幻灯片 | `document_uri`, `from_index`, `to_index` | - |
| P-21 | `ppt_goto_slide` | 跳转幻灯片 | `document_uri`, `slide_index` | - |

## 逐项验证要点

每个工具检查：

1. **存在性**：工具出现在 MCP Inspector Tools 列表中
2. **名称**：`ppt_` 前缀，snake_case 命名
3. **描述**：非空，能让用户理解工具用途
4. **`document_uri`**：必填，类型 string
5. **其他参数**：必填/可选与上表一致，类型合理
