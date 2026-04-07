---
title: "Demo Presentation"
subtitle: "pptx-skills example"
template: "templates/template.pptx"
---

## Title Page
<!-- layout: "Title Slide" -->

---

## Introduction

This is a demo slide with bullet points:

- First point with **bold** text
- Second point with *italic* text
  - A sub-bullet point
  - Another sub-bullet

---

## Data Overview

> This blockquote goes to the caption area on the left side.

| Metric | Value |
|--------|-------|
| Users  | 1,200 |
| Events | 50K   |
| Alerts | 12    |

---

## Image Slide

![Architecture diagram](architecture.png)

---

## Revenue Analysis

Quarterly revenue shows strong growth momentum.

| Quarter | Revenue |
|---------|---------|
| Q1 | $10M |
| Q2 | $15M |

```chart
type: column
title: "Quarterly Revenue"
categories: [Q1, Q2, Q3, Q4]
series:
  - name: "Revenue ($M)"
    values: [10, 15, 22, 28]
    color: "#C00000"
```

---

## Summary

- pptx-skills makes slide generation reproducible
- Markdown in, PPTX out
- Consistent styling via templates
