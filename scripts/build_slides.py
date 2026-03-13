#!/usr/bin/env python3
"""Build a pptx presentation from markdown + template.

Usage:
    python3 scripts/build_slides.py input.md [-t template.pptx] [-o output.pptx] [--img-dir slides-img]

The markdown format uses YAML front matter for metadata, ## headings for slide
titles, and --- separators between slides. See templates/template.md for format
documentation.
"""

from __future__ import annotations

import argparse
import os
import sys

from pptx import Presentation
from pptx.util import Inches, Pt, Emu

from slide_utils import (
    PresentationData, SlideData, TextElement, BlockquoteElement,
    TableElement, ImageElement, MermaidElement, Paragraph, Run,
    parse_markdown, find_layout, select_layout,
    set_text_frame, collect_text_paragraphs,
    add_table_to_slide, add_image_to_slide,
    render_all_mermaid, wps_fixup,
)


# ---------------------------------------------------------------------------
# Slide population
# ---------------------------------------------------------------------------

def populate_title_slide(slide, slide_data: SlideData, pdata: PresentationData):
    """Populate a Title Slide layout (idx 0=title, 1=subtitle)."""
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = pdata.title
        elif idx == 1:
            # Subtitle with newline support
            lines = pdata.subtitle.split('\n')
            ph.text_frame.clear()
            for i, line in enumerate(lines):
                if i == 0:
                    ph.text_frame.paragraphs[0].text = line.strip()
                else:
                    p = ph.text_frame.add_paragraph()
                    p.text = line.strip()


def populate_standard_layout(slide, slide_data: SlideData):
    """Populate a Title and Content layout (idx 0=title, 1=body)."""
    # Set title
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = slide_data.title
            break

    # Collect body content
    text_elements = [e for e in slide_data.body_elements
                     if isinstance(e, (TextElement, BlockquoteElement))]
    table_elements = [e for e in slide_data.body_elements if isinstance(e, TableElement)]
    image_elements = [e for e in slide_data.body_elements if isinstance(e, ImageElement)]

    # Fill body placeholder with text
    body_ph = None
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 1:
            body_ph = ph
            break

    if body_ph and text_elements:
        all_paras = collect_text_paragraphs(text_elements)
        if all_paras:
            set_text_frame(body_ph.text_frame, all_paras)

    # Add tables below body text or in body area
    if table_elements:
        if body_ph:
            tbl_left = body_ph.left
            tbl_top = body_ph.top + body_ph.height // 2 if text_elements else body_ph.top
            tbl_width = body_ph.width
            tbl_height = body_ph.height // 2 if text_elements else body_ph.height
        else:
            tbl_left = Emu(838200)
            tbl_top = Emu(1825625)
            tbl_width = Emu(10515600)
            tbl_height = Emu(2500000)

        for tbl_el in table_elements:
            add_table_to_slide(slide, tbl_el, tbl_left, tbl_top, tbl_width, tbl_height)

    # Add images
    if image_elements:
        if body_ph:
            img_left = body_ph.left
            img_top = body_ph.top
            img_max_w = body_ph.width
            img_max_h = body_ph.height
        else:
            img_left = Emu(838200)
            img_top = Emu(1825625)
            img_max_w = Emu(10515600)
            img_max_h = Emu(4500000)

        n_images = len(image_elements)
        for ii, img_el in enumerate(image_elements):
            # Stack images vertically
            per_h = img_max_h // n_images
            y_offset = img_top + per_h * ii
            add_image_to_slide(slide, img_el.path, img_left, y_offset, img_max_w, per_h)


def populate_caption_layout(slide, slide_data: SlideData):
    """Populate a Content with Caption layout.

    idx 0 = title, idx 2 = caption text, idx 1 = body content
    """
    bq_elements = [e for e in slide_data.body_elements if isinstance(e, BlockquoteElement)]
    text_elements = [e for e in slide_data.body_elements if isinstance(e, TextElement)]
    table_elements = [e for e in slide_data.body_elements if isinstance(e, TableElement)]
    image_elements = [e for e in slide_data.body_elements if isinstance(e, ImageElement)]

    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = slide_data.title
        elif idx == 2:
            # Caption area: use blockquote content, or first text element
            caption_paras = []
            if bq_elements:
                for bq in bq_elements:
                    caption_paras.extend(bq.paragraphs)
            elif text_elements:
                # Use first text element as caption
                caption_paras = text_elements[0].paragraphs
                text_elements = text_elements[1:]
            if caption_paras:
                set_text_frame(ph.text_frame, caption_paras)
        elif idx == 1:
            # Body area: remaining text
            remaining_paras = collect_text_paragraphs(text_elements)
            if remaining_paras:
                set_text_frame(ph.text_frame, remaining_paras)

    # Add tables
    if table_elements:
        # Position table in the body area (right side of caption layout)
        body_ph = None
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1:
                body_ph = ph
                break

        if body_ph:
            tbl_left = body_ph.left
            tbl_top = body_ph.top
            tbl_width = body_ph.width
            tbl_height = body_ph.height
        else:
            tbl_left = Emu(4800000)
            tbl_top = Emu(1825625)
            tbl_width = Emu(6500000)
            tbl_height = Emu(4000000)

        for tbl_el in table_elements:
            add_table_to_slide(slide, tbl_el, tbl_left, tbl_top, tbl_width, tbl_height)

    # Add images
    if image_elements:
        body_ph = None
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1:
                body_ph = ph
                break

        if body_ph:
            img_left = body_ph.left
            img_top = body_ph.top
            img_max_w = body_ph.width
            img_max_h = body_ph.height
        else:
            img_left = Emu(4800000)
            img_top = Emu(1825625)
            img_max_w = Emu(6500000)
            img_max_h = Emu(4000000)

        n_images = len(image_elements)
        for ii, img_el in enumerate(image_elements):
            per_h = img_max_h // n_images
            y_offset = img_top + per_h * ii
            add_image_to_slide(slide, img_el.path, img_left, y_offset, img_max_w, per_h)


def populate_title_only(slide, slide_data: SlideData):
    """Populate a Title Only layout (just title + free-form images)."""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx in (0, 4):
            # idx 4 is sometimes title in Title Only layouts
            if ph.has_text_frame:
                ph.text = slide_data.title
                break

    # Add images as free shapes
    image_elements = [e for e in slide_data.body_elements if isinstance(e, ImageElement)]
    if image_elements:
        img_left = Emu(838200)
        img_top = Emu(1825625)
        img_max_w = Emu(10515600)
        img_max_h = Emu(4500000)

        n_images = len(image_elements)
        for ii, img_el in enumerate(image_elements):
            per_h = img_max_h // n_images
            y_offset = img_top + per_h * ii
            add_image_to_slide(slide, img_el.path, img_left, y_offset, img_max_w, per_h)


# ---------------------------------------------------------------------------
# Main build pipeline
# ---------------------------------------------------------------------------

def build_presentation(pdata: PresentationData, template_path: str, output_path: str,
                       img_dir: str = "slides-img"):
    """Build a pptx from parsed markdown data and a template."""
    prs = Presentation(template_path)

    # Phase 1: Render mermaid diagrams
    render_all_mermaid(pdata.slides, img_dir)

    # Phase 2: Build title slide from front matter
    title_layout = find_layout(prs, "Title Slide", master=1)
    title_slide = prs.slides.add_slide(title_layout)
    populate_title_slide(title_slide, SlideData(), pdata)
    print(f"  Slide 1: \"{pdata.title}\" -> Title Slide")

    # Phase 3: Build content slides
    for si, slide_data in enumerate(pdata.slides):
        # Select layout for content slides
        layout = select_layout(slide_data, prs)
        slide = prs.slides.add_slide(layout)

        layout_name = layout.name
        if layout_name == "Content with Caption":
            populate_caption_layout(slide, slide_data)
        elif layout_name == "Title Only":
            populate_title_only(slide, slide_data)
        else:
            populate_standard_layout(slide, slide_data)

        print(f"  Slide {si+2}: \"{slide_data.title}\" -> {layout_name}")

    # Phase 4: Save
    prs.save(output_path)

    # Phase 5: WPS compatibility fixup
    wps_fixup(output_path)

    return len(pdata.slides) + 1  # +1 for title slide


def main():
    parser = argparse.ArgumentParser(description='Build pptx from markdown')
    parser.add_argument('input', help='Input markdown file')
    parser.add_argument('-t', '--template', default=None,
                        help='Template pptx (overrides front matter)')
    parser.add_argument('-o', '--output', default=None,
                        help='Output pptx path (overrides front matter)')
    parser.add_argument('--img-dir', default='slides-img',
                        help='Directory for mermaid renders (default: slides-img/)')
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: {args.input} not found", file=sys.stderr)
        sys.exit(1)

    print(f"==> Parsing {args.input}...")
    pdata = parse_markdown(args.input)
    print(f"  Title: {pdata.title}")
    print(f"  Slides: {len(pdata.slides)}")

    template = args.template or pdata.template
    output = args.output or pdata.output

    if not os.path.exists(template):
        print(f"Error: Template {template} not found", file=sys.stderr)
        sys.exit(1)

    print(f"==> Building slides with template: {template}")
    n_slides = build_presentation(pdata, template, output, args.img_dir)

    print(f"==> Done: {output} ({n_slides} slides)")


if __name__ == '__main__':
    main()
