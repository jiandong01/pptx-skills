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

**现有方案**

 - 手动制作 PPT
 - 风格不统一
 - 难以批量更新
 - 版本管理困难

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

## 柱形图（基础）
<!-- layout: chart -->

```chart
type: column
title: "季度用户增长"
categories: [Q1, Q2, Q3, Q4]
series:
  - name: "新增用户"
    values: [1200, 1800, 2400, 3100]
    color: "#4472C4"
  - name: "活跃用户"
    values: [800, 1400, 2000, 2800]
    color: "#70AD47"
legend: true
labels: true
```

---

## 堆叠柱形图
<!-- layout: chart -->

```chart
type: column-stacked
title: "各渠道月度收入构成（万元）"
categories: [1月, 2月, 3月, 4月, 5月, 6月]
series:
  - name: "搜索"
    values: [120, 135, 140, 155, 160, 175]
    color: "#4472C4"
  - name: "直投"
    values: [80, 90, 95, 100, 110, 120]
    color: "#C00000"
  - name: "社交"
    values: [40, 50, 55, 65, 70, 80]
    color: "#70AD47"
legend: true
labels: false
```

---

## 折线图
<!-- layout: chart -->

```chart
type: line
title: "月活用户趋势（万人）"
categories: [Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec]
series:
  - name: "2025年"
    values: [52, 55, 58, 62, 65, 70, 75, 78, 82, 85, 88, 92]
    color: "#4472C4"
  - name: "2026年"
    values: [60, 64, 70, 76, 82, 88, 95, 100, 106, 112, 118, 125]
    color: "#C00000"
legend: true
labels: false
```

---

## 面积图（堆叠）
<!-- layout: chart -->

```chart
type: area-stacked
title: "流量来源构成趋势"
categories: [Q1, Q2, Q3, Q4]
series:
  - name: "自然搜索"
    values: [350, 420, 490, 560]
    color: "#4472C4"
  - name: "付费广告"
    values: [200, 230, 270, 310]
    color: "#C00000"
  - name: "社交媒体"
    values: [100, 140, 180, 220]
    color: "#70AD47"
legend: true
labels: false
```

---

## 雷达图（能力评估）
<!-- layout: chart -->

```chart
type: radar
title: "产品能力雷达图"
categories: [功能完整性, 性能, 易用性, 安全性, 扩展性, 文档质量]
series:
  - name: "当前版本"
    values: [80, 70, 75, 85, 65, 70]
    color: "#4472C4"
  - name: "目标版本"
    values: [90, 85, 90, 90, 85, 88]
    color: "#C00000"
legend: true
labels: false
```

---

## 散点图（相关性分析）
<!-- layout: chart -->

```chart
type: scatter
title: "广告投入 vs 用户转化率"
series:
  - name: "搜索渠道"
    x_values: [10, 20, 30, 40, 50, 60]
    y_values: [0.08, 0.12, 0.15, 0.17, 0.18, 0.20]
    color: "#4472C4"
  - name: "社交渠道"
    x_values: [5, 15, 25, 35, 45]
    y_values: [0.05, 0.09, 0.11, 0.14, 0.16]
    color: "#C00000"
legend: true
```

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

## 文字 + 图表（混合布局）
<!-- layout: mixed -->

**核心业务指标分析**

- **日活用户（DAU）**：12.4万，环比 +8%
- **用户留存率（7日）**：62%，行业均值 45%
- **平均使用时长**：18.5分钟/日
- **付费转化率**：4.2%，较上季度 +0.6pp

```chart
type: column-stacked-100
title: "用户质量分层"
categories: [Q1, Q2, Q3, Q4]
series:
  - name: "高价值"
    values: [15, 18, 22, 26]
    color: "#4472C4"
  - name: "活跃"
    values: [35, 38, 40, 42]
    color: "#70AD47"
  - name: "偶发"
    values: [50, 44, 38, 32]
    color: "#A0A0A0"
legend: true
labels: false
```

---

## 总结
<!-- layout: summary -->

- **Markdown 驱动**：内容与样式分离，专注写作
- **5 类页面**：封面 / 目录 / 章节 / 内容 / 总结
- **BI 图表**：柱/条/折线/面积/饼/环/雷达/散点/气泡/瀑布/混合
- **模板兼容**：支持任意 .pptx 模板，按名称/索引/别名查找布局

