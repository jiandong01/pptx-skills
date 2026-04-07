---
title: "pptx-skills 演示"
subtitle: "Markdown → PPTX 自动化生成"
template: "template/default.pptx"
output: "examples/demo.pptx"
---

## 目录
<!-- layout: toc -->

1. 基础文本布局
2. 双栏对比布局
3. 图表布局
4. 总结

---

## 第一部分：基础文本布局
<!-- layout: section -->

标题 + 内容页，支持列表、加粗、斜体等 Markdown 格式。

---

## 功能特性
<!-- layout: standard -->

通过 Markdown 写 slide，自动渲染为 PPTX：

- 支持 **加粗** 和 *斜体* 格式
- 支持多级列表
  - 二级列表项
  - 另一个二级项
- 支持 Markdown 表格（自动居中）
- 自动根据内容类型选择布局

---

## 数据总览
<!-- layout: standard -->

| 指标 | 数值 | 环比 |
|------|------|------|
| 用户数 | 12,400 | +8% |
| 事件量 | 50万 | +15% |
| 告警数 | 12 | -25% |
| 响应时间 | 120ms | -10% |

---

## 第二部分：双栏对比布局
<!-- layout: section -->

使用 blockquote 划分左右两栏。

---

## 方案对比
<!-- layout: two-column -->

> **现有方案**
>
> - 手动制作 PPT
> - 风格不统一
> - 难以批量更新
> - 版本管理困难

**新方案**

- Markdown 编写内容
- 模板统一风格
- 批量生成，一键更新
- 可纳入版本控制

---

## 第三部分：图表布局
<!-- layout: section -->

使用 ```chart 块嵌入原生 Excel 图表。

---

## 季度增长趋势
<!-- layout: chart -->

```chart
type: column
title: "季度用户增长"
categories: [Q1, Q2, Q3, Q4]
series:
  - name: "新增用户"
    values: [1200, 1800, 2400, 3100]
    color: "#C00000"
  - name: "活跃用户"
    values: [800, 1400, 2000, 2800]
    color: "#4472C4"
labels: true
```

---

## 渠道占比
<!-- layout: chart -->

季度流量来源分析：

| 渠道 | 占比 |
|------|------|
| 搜索 | 42% |
| 直接 | 28% |
| 推荐 | 18% |
| 社交 | 12% |

```chart
type: pie
title: "流量来源分布"
categories: [搜索, 直接访问, 推荐, 社交媒体]
series:
  - name: "占比"
    values: [42, 28, 18, 12]
position: right
width: "45%"
```

---

## 总结
<!-- layout: summary -->

- **Markdown 驱动**：内容与样式分离，专注写作
- **5 类页面**：封面 / 目录 / 章节 / 内容 / 总结
- **内容自适应**：图表、双栏、表格、图片智能布局
- **模板兼容**：支持任意 .pptx 模板，按名称/索引/别名查找布局
