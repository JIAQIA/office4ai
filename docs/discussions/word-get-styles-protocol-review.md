# word:get:styles 协议一致性讨论

> **状态**: 待讨论
> **创建日期**: 2026-02-04
> **参与方**: Add-In 开发者、协议维护者、E2E 测试工程师

---

## 背景

在对 `word:get:styles` 事件进行 E2E 测试迁移时，发现当前 Add-In 实现与 OASP 协议定义存在差异。本文档旨在明确差异点，供各方讨论并达成一致。

---

## 协议来源

- **OASP 协议文档**: https://doc.turingfocus.cn/oasp/latest/specification/events-word/#wordgetstyles

---

## 差异分析

### 1. `options` 参数差异

#### OASP 协议定义

```typescript
interface GetStylesRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    includeBuiltIn?: boolean;  // 是否包含内置样式，默认 true
    includeCustom?: boolean;   // 是否包含自定义样式，默认 true
    typeFilter?: StyleType[];  // 筛选特定类型的样式
  };
}
```

#### Add-In 实际支持（基于 E2E 测试）

```typescript
options?: {
  includeBuiltIn?: boolean;   // ✅ 协议有
  includeCustom?: boolean;    // ✅ 协议有
  includeUnused?: boolean;    // ❓ 协议无，Add-In 支持
  detailedInfo?: boolean;     // ❓ 协议无，Add-In 支持
  // typeFilter?: StyleType[]; // ❓ 协议有，Add-In 是否支持？
}
```

#### 差异对比表

| 参数 | OASP 协议 | Add-In 实现 | 差异说明 |
|------|----------|-------------|---------|
| `includeBuiltIn` | ✅ 定义 | ✅ 支持 | 一致 |
| `includeCustom` | ✅ 定义 | ✅ 支持 | 一致 |
| `typeFilter` | ✅ 定义 | ❓ 待确认 | **协议有，Add-In 待确认** |
| `includeUnused` | ❌ 未定义 | ✅ 支持 | **Add-In 扩展** |
| `detailedInfo` | ❌ 未定义 | ✅ 支持 | **Add-In 扩展** |

---

### 2. 讨论问题

#### 问题 1: `includeUnused` 参数

**当前行为**: Add-In 支持 `includeUnused` 参数，用于控制是否返回文档中未使用的样式。

**讨论点**:
- [ ] 是否应将 `includeUnused` 纳入 OASP 协议？
- [ ] 如果纳入，默认值应该是什么？（建议 `false`）
- [ ] 如果不纳入，Add-In 是否应移除此参数？

**建议**: 纳入协议。该参数在实际使用中有明确需求场景（如获取文档完整样式列表）。

---

#### 问题 2: `detailedInfo` 参数

**当前行为**: Add-In 支持 `detailedInfo` 参数，用于控制是否返回样式的详细描述信息。

**讨论点**:
- [ ] 是否应将 `detailedInfo` 纳入 OASP 协议？
- [ ] `description` 字段是始终返回还是仅在 `detailedInfo=true` 时返回？
- [ ] 如果不纳入，Add-In 是否应移除此参数？

**建议**: 需要明确 `description` 字段的返回策略。如果始终返回，则不需要此参数；如果是可选的（减少数据量），则应纳入协议。

---

#### 问题 3: `typeFilter` 参数

**当前行为**: OASP 协议定义了 `typeFilter` 参数，但 E2E 测试中未覆盖。

**讨论点**:
- [ ] Add-In 是否已实现 `typeFilter` 参数？
- [ ] 如果未实现，计划何时支持？
- [ ] 如果不计划支持，是否应从协议中移除？

**建议**: 确认实现状态，并保持协议与实现一致。

---

## 响应数据一致性

### `StyleInfo` 字段定义

| 字段 | OASP 协议 | Add-In 实现 | 状态 |
|------|----------|-------------|------|
| `name` | string | string | ✅ 一致 |
| `type` | StyleType | StyleType | ✅ 一致 |
| `builtIn` | boolean | boolean | ✅ 一致 |
| `inUse` | boolean | boolean | ✅ 一致 |
| `description` | string \| null | string \| null | ✅ 一致 |

响应数据结构基本一致，无重大差异。

---

## 建议方案

### 方案 A: 更新 OASP 协议（推荐）

将 Add-In 已实现的扩展参数纳入协议：

```typescript
interface GetStylesRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    includeBuiltIn?: boolean;   // 默认 true
    includeCustom?: boolean;    // 默认 true
    includeUnused?: boolean;    // 默认 false (新增)
    detailedInfo?: boolean;     // 默认 false (新增，或移除)
    typeFilter?: StyleType[];   // 筛选特定类型
  };
}
```

**优点**: 保留现有功能，协议更完整
**缺点**: 需要更新协议文档

### 方案 B: Add-In 回退到协议定义

移除 Add-In 中的非协议参数，严格遵循 OASP 定义。

**优点**: 严格协议一致
**缺点**: 可能影响现有使用方，功能减少

### 方案 C: 协议扩展机制

在 OASP 中定义扩展机制，允许实现方添加额外参数，但明确标注为"扩展"。

---

## 待办事项

- [ ] Add-In 开发者确认 `typeFilter` 实现状态
- [ ] Add-In 开发者确认 `includeUnused` 和 `detailedInfo` 的设计意图
- [ ] 协议维护者评估是否纳入扩展参数
- [ ] 各方达成一致后，同步更新协议文档和测试用例

---

## 参与者反馈

### Add-In 开发者

> _（请在此处填写反馈）_

### 协议维护者

> _（请在此处填写反馈）_

---

## 决议记录

| 日期 | 决议内容 | 参与者 |
|------|---------|-------|
| | | |

---

**文档维护者**: E2E 测试工程师
**最后更新**: 2026-02-04
