# pptx-skills

[Agent Skills](https://agentskills.io) for building PPTX presentations from Markdown.

## Skills

| Skill | Description |
|-------|-------------|
| [extract-template](extract-template/) | Extract a reusable template from an existing `.pptx` file |
| [extract-charts](extract-charts/) | Extract chart data from a `.pptx` as reusable `chart` blocks |
| [build-slides](build-slides/) | Build a `.pptx` presentation from Markdown + template |

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Extract a template from an existing presentation

```bash
PYTHONPATH=scripts python3 scripts/extract_template.py your-slides.pptx -o templates/
```

This produces:
- `templates/template.pptx` — Clean template with masters, layouts, and theme
- `templates/template.md` — Markdown skeleton documenting available layouts

### 3. (Optional) Extract charts from an existing presentation

```bash
PYTHONPATH=scripts python3 scripts/extract_charts.py your-slides.pptx -o charts.md
```

Outputs ```` ```chart ```` YAML blocks ready to paste into your slide markdown.

### 4. Write your slides in Markdown

```markdown
---
title: "My Presentation"
subtitle: "Author Name"
template: "templates/template.pptx"
---

## First Slide

- Bullet point 1
- Bullet point 2

---

## Second Slide

| Column A | Column B |
|----------|----------|
| Data 1   | Data 2   |

---

## Third Slide

Some analysis text.

` ```chart `
type: column
title: "Revenue"
categories: [Q1, Q2, Q3, Q4]
series:
  - name: "2026"
    values: [10, 20, 30, 40]
    color: "#C00000"
` ``` `
```

### 5. Build the presentation

```bash
PYTHONPATH=scripts python3 scripts/build_slides.py slides.md
```

## Supported Content

- Text with **bold** and *italic* formatting
- Bullet lists with multiple indent levels
- Markdown tables (centered, auto-sized)
- Images (`![alt](path)`)
- Blockquotes (`> text`) mapped to caption layouts
- Mermaid diagrams (requires `npx @mermaid-js/mermaid-cli`)
- Native Excel charts (```` ```chart ```` YAML blocks) — column, bar, line, pie

## Project Structure

```
pptx-skills/
├── extract-template/
│   └── SKILL.md
├── extract-charts/
│   └── SKILL.md
├── build-slides/
│   └── SKILL.md
├── scripts/
│   ├── slide_utils.py          # Shared: data structures, parser, layout, helpers
│   ├── chart_utils.py          # Chart creation via python-pptx add_chart()
│   ├── extract_template.py     # Template extraction CLI
│   ├── extract_charts.py       # Chart extraction CLI
│   └── build_slides.py         # Slide building CLI
├── examples/
│   └── demo.md
├── requirements.txt
└── README.md
```

## License

MIT
