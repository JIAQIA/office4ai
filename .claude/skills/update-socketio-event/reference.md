# Update Socket.IO Event Reference
# 更新 Socket.IO 事件参考文档

本文档提供更新 Socket.IO 事件时的详细指南，包括变更分析模板、测试更新策略和常见场景处理。

---

## 1. 协议变更分析模板

### 1.1 变更分析文档模板

在开始更新前，建议创建变更分析记录：

```markdown
## Event Update Analysis: <event_name>
## 事件更新分析: <event_name>

### 协议版本
- 旧版本: v1.x (日期/commit)
- 新版本: v2.x (日期/commit)
- Confluence: [链接]

### 变更摘要

#### 请求参数变更
| 字段路径 | 变更类型 | 旧定义 | 新定义 | 备注 |
|---------|---------|-------|-------|-----|
| `options.newField` | 新增 | - | `string \| null` | 可选字段 |
| `format.color` | 修改 | `string` | `string \| null` | 允许空值 |
| `legacyParam` | 废弃 | `string` | - | 下版本移除 |

#### 响应数据变更
| 字段路径 | 变更类型 | 旧定义 | 新定义 | 备注 |
|---------|---------|-------|-------|-----|
| `data.metadata` | 新增 | - | `object` | 详细元数据 |

#### 行为变更
| 场景 | 旧行为 | 新行为 | 备注 |
|-----|-------|-------|-----|
| 空选区 | 返回错误 3002 | 返回空数据 | 需更新错误处理 |

### 影响评估
- [ ] DTO 需要更新
- [ ] Namespace 需要更新
- [ ] Service 需要更新
- [ ] 单元测试需要更新
- [ ] 契约测试需要更新
- [ ] 手动测试需要更新
```

### 1.2 字段变更类型判断

| 变更类型 | 判断标准 | 向后兼容 |
|---------|---------|---------|
| **新增可选** | 新字段，有默认值或可为 None | ✅ 是 |
| **新增必填** | 新字段，无默认值且不可为 None | ❌ 否 |
| **类型放宽** | `str` → `str \| None` | ✅ 是 |
| **类型收紧** | `str \| None` → `str` | ❌ 否 |
| **默认值变更** | 默认值改变 | ⚠️ 可能 |
| **字段废弃** | 标记 deprecated | ✅ 是 |
| **字段删除** | 完全移除 | ❌ 否 |

---

## 2. DTO 更新详细指南

### 2.1 新增字段

```python
from typing import Optional
from pydantic import Field

class ExistingRequest(BaseRequest):
    """现有请求类"""
    
    event_name: ClassVar[str] = "word:get:something"
    
    # 现有字段保持不变
    existing_field: str = Field(..., alias="existingField")
    
    # 新增可选字段
    new_optional_field: Optional[str] = Field(
        None, 
        alias="newOptionalField",
        description="新增的可选字段"
    )
    
    # 新增带默认值的字段
    new_field_with_default: bool = Field(
        False,
        alias="newFieldWithDefault",
        description="新增带默认值的字段"
    )
```

### 2.2 修改字段类型

```python
# 旧定义
class OldRequest(BaseRequest):
    color: str = Field(..., alias="color")

# 新定义 - 类型放宽
class NewRequest(BaseRequest):
    color: str | None = Field(None, alias="color")
```

### 2.3 废弃字段（保持兼容）

```python
from typing import Optional
from pydantic import Field
import warnings

class RequestWithDeprecation(BaseRequest):
    """包含废弃字段的请求"""
    
    # 废弃字段 - 保留但标记
    legacy_field: Optional[str] = Field(
        None,
        alias="legacyField",
        deprecated=True,
        description="[DEPRECATED] 此字段已废弃，请使用 newField"
    )
    
    # 新字段 - 替代废弃字段
    new_field: Optional[str] = Field(
        None,
        alias="newField",
        description="替代 legacyField 的新字段"
    )
    
    # 可选：添加验证器处理迁移
    @model_validator(mode='after')
    def handle_legacy_migration(self) -> Self:
        if self.legacy_field is not None and self.new_field is None:
            warnings.warn(
                "legacyField is deprecated, use newField instead",
                DeprecationWarning
            )
            self.new_field = self.legacy_field
        return self
```

### 2.4 嵌套对象更新

```python
class UpdatedOptions(SocketIOBaseModel):
    """更新后的选项对象"""
    
    # 现有字段
    include_text: bool = Field(True, alias="includeText")
    
    # 新增字段
    max_length: Optional[int] = Field(
        None, 
        alias="maxLength",
        ge=1,  # 添加约束
        description="最大返回长度"
    )

class UpdatedRequest(BaseRequest):
    """使用更新后选项的请求"""
    
    options: Optional[UpdatedOptions] = Field(None, alias="options")
```

---

## 3. 测试更新详细指南

### 3.1 单元测试更新模式

```python
# tests/unit_tests/office4ai/environment/workspace/dtos/test_word_dtos.py

import pytest
from pydantic import ValidationError
from office4ai.environment.workspace.dtos.word import UpdatedRequest

class TestUpdatedRequest:
    """UpdatedRequest 单元测试"""
    
    # ========== 现有测试保持不变（回归保护）==========
    
    def test_existing_valid_request(self):
        """测试现有有效请求 - 确保回归"""
        request = UpdatedRequest(
            requestId="req-123",
            documentUri="file:///test.docx",
            existingField="value"
        )
        assert request.existing_field == "value"
    
    # ========== 新增字段测试 ==========
    
    def test_new_field_with_value(self):
        """测试新字段有值的情况"""
        request = UpdatedRequest(
            requestId="req-123",
            documentUri="file:///test.docx",
            newOptionalField="new_value"
        )
        assert request.new_optional_field == "new_value"
    
    def test_new_field_default_none(self):
        """测试新字段默认为 None"""
        request = UpdatedRequest(
            requestId="req-123",
            documentUri="file:///test.docx"
        )
        assert request.new_optional_field is None
    
    def test_new_field_with_default_value(self):
        """测试带默认值的新字段"""
        request = UpdatedRequest(
            requestId="req-123",
            documentUri="file:///test.docx"
        )
        assert request.new_field_with_default is False
    
    # ========== 修改字段测试 ==========
    
    def test_modified_field_accepts_none(self):
        """测试修改后的字段接受 None"""
        request = UpdatedRequest(
            requestId="req-123",
            documentUri="file:///test.docx",
            color=None  # 原来不允许，现在允许
        )
        assert request.color is None
    
    # ========== 废弃字段测试 ==========
    
    def test_deprecated_field_still_works(self):
        """测试废弃字段仍然可用（向后兼容）"""
        with pytest.warns(DeprecationWarning):
            request = UpdatedRequest(
                requestId="req-123",
                documentUri="file:///test.docx",
                legacyField="old_value"
            )
        # 验证迁移逻辑
        assert request.new_field == "old_value"
    
    # ========== 边界情况测试 ==========
    
    def test_new_field_validation(self):
        """测试新字段的验证约束"""
        with pytest.raises(ValidationError):
            UpdatedRequest(
                requestId="req-123",
                documentUri="file:///test.docx",
                maxLength=-1  # 违反 ge=1 约束
            )
```

### 3.2 契约测试更新模式

```python
# tests/contract_tests/word/test_updated_event.py

import pytest
from office4ai.environment.workspace.base import OfficeAction

@pytest.mark.asyncio
@pytest.mark.contract
class TestUpdatedEventContract:
    """更新后事件的契约测试"""
    
    # ========== 现有测试保持（回归保护）==========
    
    async def test_basic_flow_unchanged(self, workspace, mock_addin_client):
        """基本流程保持不变"""
        mock_addin_client.register_static_response(
            "word:get:something",
            {"success": True, "data": {"text": "content"}}
        )
        
        action = OfficeAction(
            category="word",
            action_name="get:something",
            params={"document_uri": "file:///test.docx"}
        )
        
        result = await workspace.execute(action)
        assert result.success is True
    
    # ========== 新功能测试 ==========
    
    async def test_new_field_in_request(self, workspace, mock_addin_client):
        """测试请求中的新字段"""
        received_data = {}
        
        def capture_request(req):
            received_data.update(req)
            return {"success": True, "data": {}}
        
        mock_addin_client.register_response("word:get:something", capture_request)
        
        action = OfficeAction(
            category="word",
            action_name="get:something",
            params={
                "document_uri": "file:///test.docx",
                "new_optional_field": "test_value"
            }
        )
        
        await workspace.execute(action)
        
        # 验证新字段被正确传递
        assert received_data.get("newOptionalField") == "test_value"
    
    async def test_new_field_in_response(self, workspace, mock_addin_client):
        """测试响应中的新字段"""
        mock_addin_client.register_static_response(
            "word:get:something",
            {
                "success": True,
                "data": {
                    "text": "content",
                    "metadata": {"wordCount": 100}  # 新增响应字段
                }
            }
        )
        
        action = OfficeAction(
            category="word",
            action_name="get:something",
            params={"document_uri": "file:///test.docx"}
        )
        
        result = await workspace.execute(action)
        assert result.data.get("metadata", {}).get("wordCount") == 100
    
    # ========== 行为变更测试 ==========
    
    async def test_empty_selection_new_behavior(self, workspace, mock_addin_client):
        """测试空选区的新行为（返回空数据而非错误）"""
        mock_addin_client.register_static_response(
            "word:get:something",
            {"success": True, "data": {"text": "", "isEmpty": True}}
        )
        
        action = OfficeAction(
            category="word",
            action_name="get:something",
            params={"document_uri": "file:///test.docx"}
        )
        
        result = await workspace.execute(action)
        # 新行为：成功但数据为空
        assert result.success is True
        assert result.data.get("isEmpty") is True
```

### 3.3 手动测试更新模式

#### README.md 更新

```markdown
## 测试场景

### 现有测试（保持）

| 测试编号 | 测试名称 | 参数 | 描述 |
|---------|---------|-----|------|
| 1 | 基础获取 | 默认 | 基础功能验证 |
| 2 | 带选项获取 | `{includeText: true}` | 选项参数验证 |

### 新增测试（v2.0）

| 测试编号 | 测试名称 | 参数 | 描述 |
|---------|---------|-----|------|
| 5 | 最大长度限制 | `{maxLength: 100}` | 新增 maxLength 参数 |
| 6 | 详细元数据 | `{detailedMetadata: true}` | 新增元数据返回 |

### 废弃测试（标记删除）

| 测试编号 | 原测试名称 | 废弃原因 |
|---------|----------|---------|
| ~~3~~ | ~~旧参数测试~~ | legacyParam 已废弃 |

## 最后更新

YYYY-MM-DD - 更新至协议 v2.0
```

#### 测试文件更新

```python
# manual_tests/<event_name>_e2e/test_options_xxx.py

# 新增测试函数
async def test_5_max_length():
    """测试 5: 最大长度限制（新增）"""
    return await run_test_template(
        test_name="最大长度限制",
        test_number=5,
        action_name="get:something",
        params={"options": {"maxLength": 100}}
    )

async def test_6_detailed_metadata():
    """测试 6: 详细元数据（新增）"""
    return await run_test_template(
        test_name="详细元数据",
        test_number=6,
        action_name="get:something",
        params={"options": {"detailedMetadata": True}}
    )

# 更新 TEST_MAPPING
TEST_MAPPING = {
    "1": test_1_basic,
    "2": test_2_with_options,
    # "3": test_3_legacy,  # 已废弃，注释或删除
    "4": test_4_edge_case,
    "5": test_5_max_length,      # 新增
    "6": test_6_detailed_metadata,  # 新增
    "all": run_all_tests,
}
```

---

## 4. 常见更新场景

### 4.1 添加新的可选参数

**影响范围**: DTO, 单元测试, 手动测试

**步骤**:
1. 在 DTO 中添加 `Optional[T] = Field(None, alias="...")`
2. 添加单元测试验证字段解析
3. 在手动测试中添加新参数的测试场景

### 4.2 修改字段默认值

**影响范围**: DTO, 可能影响所有测试

**步骤**:
1. 修改 DTO 中的 `Field(default=new_value, ...)`
2. 检查所有依赖默认值的测试
3. 更新测试断言以匹配新默认值

### 4.3 废弃并替换字段

**影响范围**: DTO, 所有测试层级

**步骤**:
1. 在 DTO 中添加新字段
2. 标记旧字段为 deprecated
3. 添加迁移验证器（可选）
4. 更新测试使用新字段
5. 保留旧字段测试（验证向后兼容）

### 4.4 修改错误处理行为

**影响范围**: Namespace, 契约测试, 手动测试

**步骤**:
1. 修改 Namespace 中的错误处理逻辑
2. 更新契约测试的错误场景
3. 更新手动测试的边界情况

---

## 5. 验证检查清单

### 更新前
- [ ] 已阅读并理解协议变更文档
- [ ] 已创建变更分析记录
- [ ] 已识别所有受影响的文件

### 代码更新后
- [ ] DTO 字段定义正确
- [ ] 字段 alias 与协议一致
- [ ] 类型注解正确
- [ ] 废弃标记已添加（如适用）

### 测试更新后
- [ ] 现有单元测试仍通过
- [ ] 新增字段有单元测试
- [ ] 现有契约测试仍通过
- [ ] 新功能有契约测试
- [ ] 手动测试 README 已更新
- [ ] 手动测试用例已更新

### 最终验证
```bash
# 代码质量
poe lint
poe typecheck

# 自动化测试
poe test-unit
poe test-contract

# 手动测试（选择性运行）
uv run python manual_tests/<event_name>_e2e/test_xxx.py --test all
```

---

**最后更新**: 2026-01-13
**维护者**: JQQ <jqq1716@gmail.com>
