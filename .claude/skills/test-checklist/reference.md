# Manual Test Reference
# 手动测试参考文档

本文档为 Claude 提供手动测试用例设计的详细参考，包含参数组合矩阵、测试代码模板和常见问题排查。

---

## 1. 测试用例设计原则

### 1.1 参数组合策略

手动测试的核心价值在于覆盖各种参数排列组合，发现自动化测试难以覆盖的边界情况。

**组合优先级：**
- **P0 (必测)**: 核心功能参数的基本组合
- **P1 (重要)**: 格式参数的单项测试
- **P2 (补充)**: 多参数组合、边界值测试

**测试分组原则：**
1. 每个测试文件聚焦一个维度（如 basic、location、format）
2. 每个测试函数验证一个具体场景
3. 支持单独运行和批量运行

### 1.2 目录结构规范

```
manual_tests/<event_name>_e2e/
├── README.md              # 测试说明和运行方式
├── __init__.py            # 包初始化（可为空）
├── test_basic_xxx.py      # 基础功能测试
├── test_<dimension>.py    # 按维度分组的测试
└── test_edge_cases.py     # 边界情况测试
```

---

## 2. 各事件参数组合矩阵

### 2.1 word:insert:text

**DTO 定义：** `WordInsertTextRequest`

| 参数 | 类型 | 可选值 | 默认值 | 组合优先级 |
|-----|------|-------|-------|-----------|
| `text` | str | 短文本/长文本/多行/特殊字符/空字符串 | 必填 | P0 |
| `location` | Literal | `"Cursor"` / `"Start"` / `"End"` | `"Cursor"` | P0 |
| `format.bold` | bool \| None | `true` / `false` / `null` | `null` | P1 |
| `format.italic` | bool \| None | `true` / `false` / `null` | `null` | P1 |
| `format.underline` | bool \| None | `true` / `false` / `null` | `null` | P1 |
| `format.fontSize` | int \| None | `12` / `16` / `24` / `null` | `null` | P1 |
| `format.fontName` | str \| None | `"Arial"` / `"Times New Roman"` / `"Courier New"` / `null` | `null` | P2 |
| `format.color` | str \| None | `"#FF0000"` / `"#00FF00"` / `"#0000FF"` / `null` | `null` | P2 |

**推荐测试组合：**

```
test_basic_insert.py (4 tests):
  1. 简单文本插入: text="Hello World"
  2. 多行文本插入: text="行1\n行2\n行3"
  3. 特殊字符插入: text="@#$%^&*()_+-=[]{}|;':\",./<>?~`"
  4. 长文本插入: text=200+字符

test_location_insert.py (4 tests):
  1. 光标位置插入: location="Cursor"
  2. 文档开头插入: location="Start"
  3. 文档末尾插入: location="End"
  4. 连续多次插入: Start → End → Cursor

test_format_insert.py (6 tests):
  1. 粗体文本: format={bold: true}
  2. 斜体文本: format={italic: true}
  3. 字体大小: format={fontSize: 12/16/24}
  4. 字体名称: format={fontName: "Arial"/"Times"/"Courier"}
  5. 颜色设置: format={color: "#FF0000"/"#00FF00"/"#0000FF"}
  6. 组合格式: format={bold, italic, fontSize, fontName, color}
```

---

### 2.2 word:get:selectedContent

**DTO 定义：** `WordGetSelectedContentRequest`

| 参数 | 类型 | 可选值 | 默认值 | 组合优先级 |
|-----|------|-------|-------|-----------|
| `options.includeText` | bool | `true` / `false` | `true` | P0 |
| `options.includeImages` | bool | `true` / `false` | `true` | P1 |
| `options.includeTables` | bool | `true` / `false` | `true` | P1 |
| `options.includeContentControls` | bool | `true` / `false` | `false` | P2 |
| `options.detailedMetadata` | bool | `true` / `false` | `false` | P1 |
| `options.maxTextLength` | int \| None | `100` / `500` / `null` | `null` | P1 |

**推荐测试组合：**

```
test_basic_get.py (4 tests):
  1. 无选中内容: 光标在文档中，未选中任何内容
  2. 选中纯文本: 选中一段普通文本
  3. 选中多段落: 选中跨段落的文本
  4. 选中整个文档: Ctrl+A 全选

test_options_get.py (5 tests):
  1. 仅文本: options={includeText: true, includeImages: false, includeTables: false}
  2. 含图片: options={includeImages: true} + 选中含图片区域
  3. 含表格: options={includeTables: true} + 选中含表格区域
  4. 详细元数据: options={detailedMetadata: true}
  5. 文本截断: options={maxTextLength: 100} + 选中长文本

test_edge_cases.py (4 tests):
  1. 选中空段落: 选中只有换行的段落
  2. 选中特殊字符: 选中 emoji、数学符号等
  3. 选中跨页内容: 选中跨越分页符的内容
  4. 选中内容控件: options={includeContentControls: true}
```

---

### 2.3 word:replace:selection

**DTO 定义：** `WordReplaceSelectionRequest`

| 参数 | 类型 | 可选值 | 默认值 | 组合优先级 |
|-----|------|-------|-------|-----------|
| `content.text` | str \| None | 任意文本 / `null` | `null` | P0 |
| `content.images` | list \| None | 图片数组 / `null` | `null` | P2 |
| `content.format.bold` | bool \| None | `true` / `false` / `null` | `null` | P1 |
| `content.format.italic` | bool \| None | `true` / `false` / `null` | `null` | P1 |
| `content.format.fontSize` | int \| None | `12` / `16` / `24` / `null` | `null` | P1 |
| `content.format.fontName` | str \| None | 字体名称 / `null` | `null` | P2 |
| `content.format.color` | str \| None | 颜色值 / `null` | `null` | P2 |

**推荐测试组合：**

```
test_text_replace.py (4 tests):
  1. 纯文本替换: content={text: "新文本"}
  2. 空内容替换(删除): content={text: ""}
  3. 多行文本替换: content={text: "行1\n行2\n行3"}
  4. 特殊字符替换: content={text: "@#$%^&*"}

test_format_replace.py (4 tests):
  1. 粗体替换: content={text: "...", format: {bold: true}}
  2. 斜体替换: content={text: "...", format: {italic: true}}
  3. 颜色替换: content={text: "...", format: {color: "#FF0000"}}
  4. 组合格式替换: content={text: "...", format: {bold, italic, fontSize, color}}

test_edge_cases.py (3 tests):
  1. 无选中内容时替换: 预期失败或无操作
  2. 替换含图片区域: 选中含图片的内容后替换
  3. 替换后撤销验证: 替换后 Ctrl+Z 撤销
```

---

### 2.4 word:get:styles

**DTO 定义：** `WordGetStylesRequest`

| 参数 | 类型 | 可选值 | 默认值 | 组合优先级 |
|-----|------|-------|-------|-----------|
| `options.includeBuiltIn` | bool | `true` / `false` | `true` | P0 |
| `options.includeCustom` | bool | `true` / `false` | `true` | P0 |
| `options.includeUnused` | bool | `true` / `false` | `false` | P1 |
| `options.detailedInfo` | bool | `true` / `false` | `false` | P1 |

**推荐测试组合：**

```
test_styles.py (5 tests):
  1. 获取所有样式: options=默认
  2. 仅内置样式: options={includeBuiltIn: true, includeCustom: false}
  3. 仅自定义样式: options={includeBuiltIn: false, includeCustom: true}
  4. 包含未使用样式: options={includeUnused: true}
  5. 详细样式信息: options={detailedInfo: true}
```

---

### 2.5 word:replace:text (查找替换)

**DTO 定义：** `WordReplaceTextRequest`

| 参数 | 类型 | 可选值 | 默认值 | 组合优先级 |
|-----|------|-------|-------|-----------|
| `searchText` | str | 搜索文本 | 必填 | P0 |
| `replaceText` | str | 替换文本 | 必填 | P0 |
| `options.matchCase` | bool | `true` / `false` | `false` | P1 |
| `options.matchWholeWord` | bool | `true` / `false` | `false` | P1 |
| `options.replaceAll` | bool | `true` / `false` | `false` | P0 |

**推荐测试组合：**

```
test_replace_text.py (6 tests):
  1. 替换首个匹配: options={replaceAll: false}
  2. 替换所有匹配: options={replaceAll: true}
  3. 区分大小写: options={matchCase: true}
  4. 全词匹配: options={matchWholeWord: true}
  5. 无匹配情况: searchText 不存在于文档
  6. 组合选项: options={matchCase: true, matchWholeWord: true, replaceAll: true}
```

---

## 3. 测试代码模板

### 3.1 基础模板

```python
"""
{Event Name} E2E Test
{事件名称} 端到端测试

测试场景:
1. {场景1描述}
2. {场景2描述}
...
"""

import asyncio
import sys
from contextlib import asynccontextmanager

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@asynccontextmanager
async def workspace_context(host: str = "127.0.0.1", port: int = 3000):
    """
    Workspace 上下文管理器，自动处理启动和停止
    Context manager for Workspace, handles start and stop automatically
    """
    workspace = OfficeWorkspace(host=host, port=port)
    try:
        await workspace.start()
        yield workspace
    finally:
        await workspace.stop()


async def wait_for_connection(workspace: OfficeWorkspace, timeout: float = 30.0) -> bool:
    """
    等待 Add-In 连接
    Wait for Add-In connection
    """
    print("\n⏳ 等待 Word Add-In 连接...")
    connected = await workspace.wait_for_addin_connection(timeout=timeout)
    if not connected:
        print("❌ 超时：未检测到 Add-In 连接")
        return False
    return True


def get_document_uri(workspace: OfficeWorkspace) -> str | None:
    """
    获取已连接文档的 URI
    Get connected document URI
    """
    documents = workspace.get_connected_documents()
    if not documents:
        print("❌ 未找到已连接文档")
        return None
    return documents[0]


async def run_test_template(
    test_name: str,
    test_number: int,
    action_name: str,
    params: dict,
    wait_seconds: int = 3,
) -> bool:
    """
    测试执行模板：封装通用的测试流程
    Test execution template: encapsulates common test flow
    """
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_name}")
    print("=" * 70)

    async with workspace_context() as workspace:
        try:
            print("✅ Workspace 启动成功")

            if not await wait_for_connection(workspace):
                return False

            document_uri = get_document_uri(workspace)
            if not document_uri:
                return False

            print(f"✅ 使用文档: {document_uri}")
            print(f"\n📝 执行动作: {action_name}")
            print(f"   参数: {params}")

            await asyncio.sleep(wait_seconds)

            action = OfficeAction(
                category="word",
                action_name=action_name,
                params={"document_uri": document_uri, **params},
            )

            result = await workspace.execute(action)

            print("\n📊 验证结果:")
            if result.success:
                print("✅ 执行成功")
                print(f"   返回数据: {result.data}")
                return True
            else:
                print(f"❌ 执行失败: {result.error}")
                return False

        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback
            traceback.print_exc()
            return False


# 测试映射表
TEST_MAPPING = {
    "1": lambda: run_test_template("测试1名称", 1, "action:name", {"param": "value"}),
    "2": lambda: run_test_template("测试2名称", 2, "action:name", {"param": "value"}),
    "all": run_all_tests,
}


async def run_all_tests():
    """运行所有测试"""
    print("\n🚀 运行所有测试...\n")
    results = []
    for key, test_func in TEST_MAPPING.items():
        if key != "all":
            results.append(await test_func())
            await asyncio.sleep(2)
    
    print("\n" + "=" * 70)
    print(f"📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    print("=" * 70)
    return all(results)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="E2E Tests")
    parser.add_argument(
        "--test",
        choices=list(TEST_MAPPING.keys()),
        default="1",
        help="Test to run",
    )

    args = parser.parse_args()

    try:
        test_func = TEST_MAPPING[args.test]
        success = asyncio.run(test_func())
        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)
```

### 3.2 获取内容测试模板

```python
async def test_get_content(test_number: int, test_name: str, options: dict) -> bool:
    """
    获取内容测试模板
    Get content test template
    """
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_name}")
    print("=" * 70)

    async with workspace_context() as workspace:
        try:
            print("✅ Workspace 启动成功")

            if not await wait_for_connection(workspace):
                return False

            document_uri = get_document_uri(workspace)
            if not document_uri:
                return False

            print(f"✅ 使用文档: {document_uri}")
            print(f"\n📝 获取选中内容")
            print(f"   选项: {options}")
            print("\n   ⚠️  请在 Word 中选中一些内容...")

            await asyncio.sleep(5)  # 给用户时间选中内容

            action = OfficeAction(
                category="word",
                action_name="get:selectedContent",
                params={"document_uri": document_uri, "options": options},
            )

            result = await workspace.execute(action)

            print("\n📊 验证结果:")
            if result.success:
                print("✅ 获取成功")
                data = result.data
                if "text" in data:
                    text = data["text"]
                    print(f"   文本内容: '{text[:100]}{'...' if len(text) > 100 else ''}'")
                    print(f"   文本长度: {len(text)}")
                if "images" in data:
                    print(f"   图片数量: {len(data['images'])}")
                if "tables" in data:
                    print(f"   表格数量: {len(data['tables'])}")
                return True
            else:
                print(f"❌ 获取失败: {result.error}")
                return False

        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback
            traceback.print_exc()
            return False
```

---

## 4. 常见问题排查

### 4.1 连接问题

| 问题 | 症状 | 排查步骤 |
|-----|------|---------|
| Add-In 无法连接 | 超时等待连接 | 1. 检查 Workspace 是否启动 `curl http://127.0.0.1:3000/health`<br>2. 检查 Add-In 控制台错误<br>3. 检查防火墙设置 |
| 文档未找到 | `❌ 未找到已连接文档` | 1. 确认 Word 文档已打开<br>2. 确认 Add-In 已加载<br>3. 重新加载 Add-In |
| 请求超时 | `Timeout waiting for response` | 1. 检查 Add-In 是否响应<br>2. 查看 Add-In 控制台日志<br>3. 增加超时时间 |

### 4.2 操作问题

| 问题 | 症状 | 排查步骤 |
|-----|------|---------|
| 插入位置不对 | 文本未插入到预期位置 | 1. 确认 location 参数正确<br>2. 对于 Cursor，确认光标位置<br>3. 检查文档是否只读 |
| 格式未应用 | 插入的文本无格式 | 1. 检查 format 参数格式<br>2. 确认 Add-In 支持该格式<br>3. 在 Word 中手动检查 |
| 获取内容为空 | 返回空数据 | 1. 确认已选中内容<br>2. 检查 options 参数<br>3. 尝试选中更多内容 |

### 4.3 调试技巧

1. **查看 Workspace 日志**
   ```bash
   # 启动时会输出详细日志
   uv run python manual_tests/xxx.py --test 1
   ```

2. **查看 Add-In 控制台**
   - 在 Word 中按 F12 打开开发者工具
   - 查看 Console 和 Network 标签

3. **手动测试 Socket.IO**
   ```javascript
   // 在 Add-In 控制台执行
   const client = window.socketClient;
   console.log('Connected:', client.isConnected());
   console.log('Document URI:', client.getDocumentUri());
   ```

---

## 5. 测试报告模板

```markdown
## Manual Test Report
## 手动测试报告

**测试日期**: YYYY-MM-DD
**测试人员**: [Name]
**测试环境**:
- office4ai commit: [hash]
- office-editor4ai commit: [hash]
- Python: [version]
- Node.js: [version]
- Word: [version]

### 测试结果

| 测试套件 | 通过 | 失败 | 跳过 | 备注 |
|---------|-----|-----|-----|------|
| insert_text_e2e | x/14 | x | x | |
| get_selected_content_e2e | x/13 | x | x | |
| replace_selection_e2e | x/11 | x | x | |
| get_styles_e2e | x/5 | x | x | |
| connection_e2e | x/4 | x | x | |

### 失败用例详情

| 测试 | 错误信息 | 截图 |
|-----|---------|-----|
| xxx | xxx | [link] |

### 结论

- ✅ 全部通过
- ⚠️ 部分通过，需关注
- ❌ 存在阻塞问题
```

---

**最后更新**: 2026-01-09
**维护者**: JQQ <jqq1716@gmail.com>
