# pptx-skills

[Agent Skills](https://agentskills.io) for building PPTX presentations from Markdown.

## Skills

| Skill | Description |
|-------|-------------|
| [extract-template](extract-template/) | Extract a reusable template from an existing `.pptx` file |
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

### 3. Write your slides in Markdown

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
```

### 4. Build the presentation

```bash
PYTHONPATH=scripts python3 scripts/build_slides.py slides.md
```

## Supported Content

- Text with **bold** and *italic* formatting
- Bullet lists with multiple indent levels
- Markdown tables
- Images (`![alt](path)`)
- Blockquotes (`> text`) mapped to caption layouts
- Mermaid diagrams (requires `npx @mermaid-js/mermaid-cli`)

## Project Structure

```
pptx-skills/
├── extract-template/
│   └── SKILL.md
├── build-slides/
│   └── SKILL.md
├── scripts/
│   ├── slide_utils.py
│   ├── extract_template.py
│   └── build_slides.py
├── examples/
│   └── demo.md
├── requirements.txt
└── README.md
```

## License

MIT
