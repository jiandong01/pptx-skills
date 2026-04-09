---
title: "pptx-skills 演示"
subtitle: "Markdown → PPTX 自动化生成"
template: "default.pptx"
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

文本与表格布局示例

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

左右对比，blockquote 为左栏

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

原生 Excel 图表嵌入示例

---

## 季度增长趋势（柱形 + 折线混合）
<!-- layout: chart -->

```chart
type: combo
title: "季度用户增长与增长率"
categories: [Q1, Q2, Q3, Q4]
series:
  - name: "新增用户"
    values: [1200, 1800, 2400, 3100]
    color: "#4472C4"
    subtype: bar
  - name: "活跃用户"
    values: [800, 1400, 2000, 2800]
    color: "#70AD47"
    subtype: bar
  - name: "增长率"
    values: [0.05, 0.50, 0.33, 0.29]
    color: "#C00000"
    subtype: line
    axis: secondary
    number_format: "0%"
legend: true
```

---

## 收入瀑布分析
<!-- layout: chart -->

```chart
type: waterfall
title: "Q4 利润拆解（万元）"
categories: [期初利润, 收入增长, 成本优化, 费用上升, 净利润]
series:
  - name: "金额"
    values: [0, 500, 150, -200, 450]
totals: [0, 4]
colors:
  gain: "#70AD47"
  loss: "#FF4040"
  total: "#4472C4"
labels: true
```

---

## 渠道占比（环形图）
<!-- layout: chart -->

```chart
type: doughnut
title: "流量来源分布"
categories: [搜索, 直接访问, 推荐, 社交媒体]
series:
  - name: "占比"
    values: [42, 28, 18, 12]
labels: true
legend: true
```

---

## 用户规模气泡图
<!-- layout: chart -->

```chart
type: bubble
title: "产品矩阵：用户数 vs 留存率 vs 收入规模"
series:
  - name: "产品A"
    x_values: [80, 60, 40, 70]
    y_values: [0.75, 0.60, 0.85, 0.50]
    sizes: [30, 20, 15, 25]
    color: "#4472C4"
  - name: "产品B"
    x_values: [30, 50, 90]
    y_values: [0.40, 0.65, 0.70]
    sizes: [10, 18, 35]
    color: "#C00000"
legend: true
```

---

## 总结
<!-- layout: summary -->

- **Markdown 驱动**：内容与样式分离，专注写作
- **5 类页面**：封面 / 目录 / 章节 / 内容 / 总结
- **BI 图表**：柱/条/折线/面积/饼/环/雷达/散点/气泡/瀑布/混合
- **模板兼容**：支持任意 .pptx 模板，按名称/索引/别名查找布局

