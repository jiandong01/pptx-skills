---
name: build-slides
description: Build a PPTX presentation from Markdown content using an extracted template. Supports text, bullet lists, tables, images, blockquotes, and mermaid diagrams. Use when the user wants to generate slides from markdown or create a presentation.
license: MIT
compatibility: Requires python-pptx, PyYAML, lxml, Pillow. Optional mermaid-cli (npx @mermaid-js/mermaid-cli) for diagram rendering.
metadata:
  author: pptx-skills
  version: "1.0"
---

# Build PPTX from Markdown

Build a polished PPTX presentation from a Markdown file and a template extracted by the `extract-template` skill.

## When to Use

- User wants to generate a pptx from markdown (e.g., "build slides", "生成幻灯片")
- User asks to create or update a presentation
- User has a markdown file with slide content and a template.pptx

## Instructions

**Steps:**

1. Ensure a template exists. If not, suggest running `extract-template` first.
2. Help the user write or review their markdown file (see format below).
3. Run the build script:
   ```
   PYTHONPATH=<skill_root>/scripts python3 <skill_root>/scripts/build_slides.py <input.md> [-t template.pptx] [-o output.pptx] [--img-dir slides-img]
   ```
   Replace `<skill_root>` with the absolute path to the pptx-skills installation directory.
4. Report the number of slides generated and output file path.

## Markdown Format

### Front Matter (YAML)

```yaml
---
title: "Presentation Title"
subtitle: "Subtitle text"
template: "templates/template.pptx"
---
```

- `title` / `subtitle`: Shown on the title slide. Subtitle supports `\n` for newlines.
- `template`: Path to the template pptx (from `extract-template`).
- `output`: (Optional) Output path. Defaults to `<input-name>.pptx`.

### Slide Syntax

- `## Heading` — Slide title
- `---` — Slide separator
- `<!-- layout: "Layout Name" -->` — Force a specific layout

### Supported Content

| Element | Syntax | Notes |
|---------|--------|-------|
| Bullet list | `- item` / `  - sub-item` | Indent with 2 spaces per level |
| Bold/Italic | `**bold**` / `*italic*` | Inline formatting |
| Table | Standard markdown table | Header + separator + rows |
| Image | `![alt](path)` | Auto-scaled to fit |
| Blockquote | `> text` | Mapped to caption area in "Content with Caption" layout |
| Mermaid | `` ```mermaid `` code block | Rendered to PNG via mermaid-cli |

### Layout Auto-Selection

The script automatically selects layouts based on content:

| Content Pattern | Layout Selected |
|----------------|----------------|
| Blockquote + text/table | Content with Caption |
| Text + table | Content with Caption |
| Image only (no text) | Title Only |
| Empty body | Title Only |
| Default | Title and Content |

You can override auto-selection with `<!-- layout: "Layout Name" -->`.

## Example

See [examples/demo.md](../examples/demo.md) for a complete example.
