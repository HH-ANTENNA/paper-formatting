#!/usr/bin/env python3
"""
format_docx.py — Format a Word document according to paper template rules.

Phases: A=Page+Body, B=FrontMatter, C=Headings, D=Tables+Figures, E=References
Run with --phase all (default) or specify individual phases.
"""
import json
import sys
import os
import re
from copy import deepcopy
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement


# ═══════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════

def is_chinese_char(ch):
    """Check if a character is in the CJK range."""
    cp = ord(ch)
    return (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or
            0x20000 <= cp <= 0x2A6DF or 0xF900 <= cp <= 0xFAFF or
            0x3000 <= cp <= 0x303F or 0xFF00 <= cp <= 0xFFEF)


def text_is_chinese(text):
    """Check if majority of text is Chinese."""
    if not text:
        return False
    chinese_count = sum(1 for c in text if is_chinese_char(c))
    return chinese_count > len(text) * 0.3


def set_run_font(run, chinese_font, english_font, size_pt, bold=None):
    """Set font properties for a run with CJK support."""
    run.font.size = Pt(size_pt)
    run.font.name = english_font
    if bold is not None:
        run.bold = bold

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), chinese_font)
    rFonts.set(qn('w:ascii'), english_font)
    rFonts.set(qn('w:hAnsi'), english_font)
    rFonts.set(qn('w:cs'), english_font)


def set_paragraph_spacing(paragraph, line_spacing=1.5, space_before=0, space_after=0,
                          first_line_indent_cm=0.74, alignment=None):
    """Set paragraph spacing and indentation."""
    pf = paragraph.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_line_indent_cm is not None:
        pf.first_line_indent = Cm(first_line_indent_cm)
    if alignment is not None:
        paragraph.alignment = alignment


def fix_chinese_punctuation(text):
    """Fix punctuation in Chinese text per template rules."""
    # Full stop: "。" or "." followed by space → "．"
    text = text.replace('。', '．')
    # English dot followed by space → fullwidth full stop if in Chinese context
    # (this needs context awareness, simplified here)
    # Comma: "," → "，" when in Chinese context
    # Colon: ":" → "：" when in Chinese context
    text = re.sub(r'(?<=[一-鿿]),', '，', text)
    text = re.sub(r'(?<=[一-鿿]):', '：', text)
    text = re.sub(r'(?<=[一-鿿]);', '；', text)
    return text


def detect_paragraph_type(paragraph):
    """Heuristically detect what type of content a paragraph contains."""
    text = paragraph.text.strip()
    if not text:
        return "empty"

    # Check for keywords patterns
    if re.match(r'^关键词[：:]', text):
        return "keywords_chinese"
    if re.match(r'^Key words?[：:]', text, re.IGNORECASE):
        return "keywords_english"

    # Check for abstract patterns
    if re.match(r'^摘\s*要[：:]?', text):
        return "abstract_chinese_heading"
    if re.match(r'^Abstract[：:]?', text, re.IGNORECASE):
        return "abstract_english_heading"

    # Check for headings
    if re.match(r'^第?\d+(\.\d+)*\s', text):
        # Numbered heading like "1 XXXX" or "1.1 XXXX" or "1.1.1 XXXX"
        depth = text.split()[0].count('.') + 1
        if depth == 1:
            return "heading_l1"
        elif depth == 2:
            return "heading_l2"
        else:
            return "heading_l3"

    # Check unnumbered headings
    if text in ['引言', '绪论', '结语', '结论', '致谢', '参考文献', 'References']:
        return "unnumbered_heading"

    # Check for table/figure captions
    if re.match(r'^表\s*\d+', text):
        return "table_caption_chinese"
    if re.match(r'^Table\s+\d+', text, re.IGNORECASE):
        return "table_caption_english"
    if re.match(r'^图\s*\d+', text):
        return "figure_caption_chinese"
    if re.match(r'^Fig\.?\s*\d+', text, re.IGNORECASE):
        return "figure_caption_english"

    # Check for reference entries
    if re.match(r'^\[\d+\]', text):
        return "reference_entry"

    # Check for author line (typically after title, before affiliation)
    if re.match(r'^作者\d+', text):
        return "author_chinese"
    if re.match(r'^Author\s+\d+', text, re.IGNORECASE):
        return "author_english"

    # Check for affiliation
    if re.match(r'^\d+\)', text):
        return "affiliation"

    return "body"


# ═══════════════════════════════════════════════════════════
# Phase A: Page Setup + Body Font + Paragraph + Punctuation
# ═══════════════════════════════════════════════════════════

def phase_a_page_and_body(doc, rules):
    """Apply page setup, body font, paragraph formatting, and punctuation."""
    changes = []

    # Page setup
    page_rules = rules.get("extracted", {}).get("page", rules.get("page", {}))
    for section in doc.sections:
        if page_rules.get("width_cm"):
            section.page_width = Cm(page_rules["width_cm"])
            section.page_height = Cm(page_rules["height_cm"])
            changes.append(f"Page size: {page_rules['width_cm']}x{page_rules['height_cm']}cm")
        if page_rules.get("top_margin_cm"):
            section.top_margin = Cm(page_rules["top_margin_cm"])
            section.bottom_margin = Cm(page_rules["bottom_margin_cm"])
            section.left_margin = Cm(page_rules["left_margin_cm"])
            section.right_margin = Cm(page_rules["right_margin_cm"])
            changes.append(f"Margins: T={page_rules['top_margin_cm']} B={page_rules['bottom_margin_cm']} "
                          f"L={page_rules['left_margin_cm']} R={page_rules['right_margin_cm']}cm")

    # Body font
    body = rules.get("body_font", {})
    chinese_font = body.get("chinese", "宋体")
    english_font = body.get("english", "Times New Roman")
    size_pt = body.get("size_pt", 10.5)
    line_spacing = rules.get("line_spacing", 1.5)
    indent_cm = rules.get("first_line_indent_cm", 0.74)

    # Apply to all "Normal" style paragraphs
    for p in doc.paragraphs:
        ptype = detect_paragraph_type(p)
        if ptype in ("empty",):
            continue

        text = p.text.strip()
        if not text:
            continue

        # Fix punctuation in Chinese text
        if text_is_chinese(text):
            new_text = fix_chinese_punctuation(text)
            if new_text != text:
                # Only modify if there are actual runs to update
                if p.runs:
                    for run in p.runs:
                        run.text = fix_chinese_punctuation(run.text)
                    changes.append(f"Punctuation fixed in: {text[:40]}...")

        # Apply body formatting to body paragraphs
        if ptype == "body":
            for run in p.runs:
                set_run_font(run, chinese_font, english_font, size_pt)
            set_paragraph_spacing(p, line_spacing, first_line_indent_cm=indent_cm,
                                  alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
        elif ptype == "empty":
            continue

    changes.append(f"Body font: {chinese_font}/{english_font} {size_pt}pt, "
                   f"line spacing {line_spacing}, indent {indent_cm}cm")
    return changes


# ═══════════════════════════════════════════════════════════
# Phase B: Front Matter (Title, Authors, Abstract, Keywords)
# ═══════════════════════════════════════════════════════════

def phase_b_front_matter(doc, rules):
    """Format front matter: title, authors, abstract, keywords."""
    changes = []

    title_font = rules.get("title", {"chinese": "宋体", "english": "Times New Roman", "size_pt": 22})
    author_font = rules.get("author", {"chinese": "楷体", "english": "Times New Roman", "size_pt": 14})
    abstract_font = rules.get("abstract", {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5})
    kw_font = rules.get("keywords", {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5})

    for p in doc.paragraphs:
        ptype = detect_paragraph_type(p)
        text = p.text.strip()
        if not text:
            continue

        if ptype == "empty":
            continue

        # First non-empty paragraph is typically the title
        # (We rely on position for title detection since it may not have a clear marker)
        elif ptype == "body":
            # Could be title if it's the first few paragraphs
            pass  # Title detection needs positional context

        elif ptype == "author_chinese" or ptype == "author_english":
            for run in p.runs:
                set_run_font(run, author_font["chinese"], author_font["english"],
                            author_font["size_pt"])
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(p, 1.5, first_line_indent_cm=0)
            changes.append(f"Author formatted: {text[:50]}")

        elif ptype == "abstract_chinese_heading":
            for run in p.runs:
                set_run_font(run, abstract_font["chinese"], abstract_font["english"],
                            abstract_font["size_pt"], bold=True)
            changes.append(f"Abstract heading formatted: {text[:50]}")

        elif ptype == "abstract_english_heading":
            for run in p.runs:
                set_run_font(run, abstract_font["english"], abstract_font["english"],
                            abstract_font["size_pt"], bold=True)
            changes.append(f"English abstract heading formatted: {text[:50]}")

        elif ptype == "keywords_chinese":
            for run in p.runs:
                set_run_font(run, kw_font["chinese"], kw_font["english"],
                            kw_font["size_pt"])
            # Fix keyword separator
            if p.runs:
                for run in p.runs:
                    run.text = run.text.replace(',', '；').replace('，', '；')
            changes.append(f"Keywords formatted: {text[:50]}")

        elif ptype == "keywords_english":
            for run in p.runs:
                set_run_font(run, kw_font["english"], kw_font["english"],
                            kw_font["size_pt"])
            changes.append(f"English keywords formatted: {text[:50]}")

    return changes


# ═══════════════════════════════════════════════════════════
# Phase C: Heading Hierarchy
# ═══════════════════════════════════════════════════════════

def phase_c_headings(doc, rules):
    """Format heading hierarchy: L1, L2, L3, unnumbered headings."""
    changes = []

    l1 = rules.get("heading_l1", {"chinese": "仿宋", "english": "Times New Roman", "size_pt": 16, "bold": False})
    l2 = rules.get("heading_l2", {"chinese": "黑体", "english": "Times New Roman", "size_pt": 10.5, "bold": True})
    l3 = rules.get("heading_l3", {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5, "bold": False})

    heading_map = {
        "heading_l1": l1,
        "heading_l2": l2,
        "heading_l3": l3,
        "unnumbered_heading": l1,  # Unnumbered headings use L1 style
    }

    for p in doc.paragraphs:
        ptype = detect_paragraph_type(p)
        if ptype not in heading_map:
            continue

        font_config = heading_map[ptype]
        for run in p.runs:
            set_run_font(run, font_config["chinese"], font_config["english"],
                        font_config["size_pt"], bold=font_config.get("bold"))
        set_paragraph_spacing(p, 1.5, space_before=6, space_after=3,
                              first_line_indent_cm=0, alignment=WD_ALIGN_PARAGRAPH.LEFT)
        changes.append(f"{ptype} formatted: {p.text.strip()[:50]}")

    return changes


# ═══════════════════════════════════════════════════════════
# Phase D: Tables and Figures
# ═══════════════════════════════════════════════════════════

def phase_d_tables_figures(doc, rules):
    """Apply three-line table style and format captions."""
    changes = []
    caption_font = rules.get("caption", {"chinese": "宋体", "english": "Times New Roman", "size_pt": 9})

    # Format table captions
    for p in doc.paragraphs:
        ptype = detect_paragraph_type(p)
        if ptype in ("table_caption_chinese", "table_caption_english",
                     "figure_caption_chinese", "figure_caption_english"):
            for run in p.runs:
                is_bold = "表" in run.text or "Table" in run.text or "图" in run.text or "Fig" in run.text
                # Bold the label part only (simplified)
                set_run_font(run, caption_font["chinese"], caption_font["english"],
                            caption_font["size_pt"])
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(p, 1.0, first_line_indent_cm=0)
            changes.append(f"Caption formatted: {p.text.strip()[:50]}")

    # Apply three-line table style
    for table in doc.tables:
        apply_three_line_table(table)
        # Format cell text
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        set_run_font(run, caption_font["chinese"], caption_font["english"],
                                    caption_font["size_pt"])
        changes.append(f"Table converted to three-line ({len(table.rows)} rows)")

    return changes


def apply_three_line_table(table):
    """Convert a table to three-line academic style.

    Three-line means:
    - Thick top border on header row (1.5pt)
    - Thin bottom border on header row (0.75pt)
    - Thick bottom border on last row (1.5pt)
    - No left/right/vertical borders
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    # Remove table-level borders
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil')
        tblBorders.append(border)
    # Remove existing tblBorders if present
    existing = tblPr.findall(qn('w:tblBorders'))
    for e in existing:
        tblPr.remove(e)
    tblPr.append(tblBorders)

    # Style each row
    num_rows = len(table.rows)
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')

            if ri == 0:
                # Header row: thick top + thin bottom
                top = OxmlElement('w:top')
                top.set(qn('w:val'), 'single')
                top.set(qn('w:sz'), '12')  # 1.5pt
                top.set(qn('w:color'), '000000')
                tcBorders.append(top)

                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'single')
                bottom.set(qn('w:sz'), '6')  # 0.75pt
                bottom.set(qn('w:color'), '000000')
                tcBorders.append(bottom)

            elif ri == num_rows - 1:
                # Last row: thick bottom
                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'single')
                bottom.set(qn('w:sz'), '12')
                bottom.set(qn('w:color'), '000000')
                tcBorders.append(bottom)

            # Remove left/right borders for all cells
            for side in ['left', 'right']:
                border = OxmlElement(f'w:{side}')
                border.set(qn('w:val'), 'nil')
                tcBorders.append(border)

            # Replace existing borders
            existing_borders = tcPr.findall(qn('w:tcBorders'))
            for eb in existing_borders:
                tcPr.remove(eb)
            tcPr.append(tcBorders)


# ═══════════════════════════════════════════════════════════
# Phase E: References
# ═══════════════════════════════════════════════════════════

def phase_e_references(doc, rules):
    """Format references section."""
    changes = []

    ref_title_font = rules.get("reference_title", {"chinese": "黑体", "english": "Times New Roman", "size_pt": 10.5})
    ref_body_font = rules.get("reference_body", {"chinese": "宋体", "english": "Times New Roman", "size_pt": 9})

    in_refs = False
    for p in doc.paragraphs:
        ptype = detect_paragraph_type(p)
        text = p.text.strip()
        if not text:
            continue

        # Detect reference section header
        if text in ['参考文献', 'References'] or text.startswith('参考文献'):
            in_refs = True
            for run in p.runs:
                set_run_font(run, ref_title_font["chinese"], ref_title_font["english"],
                            ref_title_font["size_pt"], bold=True)
            set_paragraph_spacing(p, 1.5, space_before=12, first_line_indent_cm=0)
            changes.append("Reference title formatted")
            continue

        if in_refs and ptype == "reference_entry":
            for run in p.runs:
                set_run_font(run, ref_body_font["chinese"], ref_body_font["english"],
                            ref_body_font["size_pt"])
            # Hanging indent for references
            pf = p.paragraph_format
            pf.first_line_indent = Cm(-0.74)
            pf.left_indent = Cm(0.74)
            changes.append(f"Reference entry formatted: {text[:50]}")

        # Fix non-standard reference numbering (e.g., "1." → "[1]")
        elif in_refs and re.match(r'^\d+\.', text):
            for run in p.runs:
                set_run_font(run, ref_body_font["chinese"], ref_body_font["english"],
                            ref_body_font["size_pt"])
            if p.runs:
                p.runs[0].text = re.sub(r'^(\d+)\.', r'[\1]', p.runs[0].text)
            changes.append(f"Reference numbering fixed: {text[:50]}")

    return changes


# ═══════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════

PHASES = {
    "A": ("Page + Body", phase_a_page_and_body),
    "B": ("Front Matter", phase_b_front_matter),
    "C": ("Headings", phase_c_headings),
    "D": ("Tables + Figures", phase_d_tables_figures),
    "E": ("References", phase_e_references),
}


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Format academic paper (.docx)")
    parser.add_argument("input", help="Input .docx file")
    parser.add_argument("rules", help="Rules JSON file (from extract_rules.py)")
    parser.add_argument("--output", "-o", required=True, help="Output .docx file")
    parser.add_argument("--phase", default="all",
                       help="Phases to run: A,B,C,D,E or 'all' (default)")
    args = parser.parse_args()

    # Load rules
    with open(args.rules, 'r', encoding='utf-8') as f:
        rules = json.load(f)

    # Load document
    doc = Document(args.input)

    # Determine phases to run
    if args.phase == "all":
        phases_to_run = ["A", "B", "C", "D", "E"]
    else:
        phases_to_run = [p.strip() for p in args.phase.split(",")]

    # Run phases
    all_changes = []
    for phase_id in phases_to_run:
        if phase_id not in PHASES:
            print(f"Unknown phase: {phase_id}")
            continue
        name, func = PHASES[phase_id]
        print(f"\n{'='*60}")
        print(f"Phase {phase_id}: {name}")
        print(f"{'='*60}")
        changes = func(doc, rules)
        for c in changes:
            print(f"  [OK] {c}")
        all_changes.extend(changes)

    # Save
    doc.save(args.output)
    print(f"\n{'='*60}")
    print(f"Formatted document saved to: {args.output}")
    print(f"Total changes: {len(all_changes)}")


if __name__ == '__main__':
    main()
