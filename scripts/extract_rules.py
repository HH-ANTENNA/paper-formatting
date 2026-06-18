#!/usr/bin/env python3
"""
extract_rules.py — Extract formatting rules from a paper template .docx

Reads a Word template and produces a JSON rules file containing both:
  1. Text rules (parsed from the template's instructional content)
  2. Visual formatting (observed font sizes, margins, etc.)

The JSON output drives the formatting scripts.
"""
import json
import sys
import os
from docx import Document
from docx.shared import Cm, Pt, Emu


def emu_to_cm(emu):
    """Convert EMU (English Metric Units) to cm."""
    if emu is None:
        return None
    return round(emu / 360000, 2)


def emu_to_pt(emu):
    """Convert EMU to points (1pt = 12700 EMU)."""
    if emu is None:
        return None
    return round(emu / 12700, 1)


def extract_rules(template_path):
    """Extract all formatting rules from the template."""
    doc = Document(template_path)
    rules = {
        "template_path": template_path,
        "page": {},
        "fonts": {},
        "headings": {},
        "paragraph": {},
        "punctuation": {},
        "table": {},
        "reference": {},
        "bilingual": {},
        "text_rules": [],  # Rules extracted from template text content
    }

    # ── Page setup ──
    section = doc.sections[0]
    rules["page"] = {
        "width_cm": emu_to_cm(section.page_width),
        "height_cm": emu_to_cm(section.page_height),
        "top_margin_cm": emu_to_cm(section.top_margin),
        "bottom_margin_cm": emu_to_cm(section.bottom_margin),
        "left_margin_cm": emu_to_cm(section.left_margin),
        "right_margin_cm": emu_to_cm(section.right_margin),
    }

    # ── Scan all paragraphs for styling and rules ──
    font_samples = {}  # element_type -> {font, size, bold}

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        style_name = p.style.name

        # Collect font samples from runs
        for run in p.runs:
            if run.font.name and run.font.size:
                key = f"{style_name}_{run.font.name}"
                if key not in font_samples:
                    font_samples[key] = {
                        "style": style_name,
                        "font": run.font.name,
                        "size_pt": emu_to_pt(run.font.size),
                        "bold": run.font.bold,
                    }

        # ── Parse text rules from the template's instructional content ──
        # Body text rule
        if "正文叙述用" in text and "号" in text:
            rules["text_rules"].append({"type": "body_font", "text": text})

        # Title rule
        if "中文题目" in text and "字" in text:
            rules["text_rules"].append({"type": "chinese_title", "text": text})

        # English title rule
        if "英文标题" in text or "英文题目" in text:
            rules["text_rules"].append({"type": "english_title", "text": text})

        # Abstract rule
        if "摘" in text and "要" in text and ("第三人称" in text or "5号楷体" in text or "撰写" in text):
            rules["text_rules"].append({"type": "abstract", "text": text})

        # Keywords rule
        if "关键词" in text and "6" in text and "12" in text:
            rules["text_rules"].append({"type": "keywords", "text": text})

        # Heading rules
        if "1级标题" in text:
            rules["text_rules"].append({"type": "heading_l1", "text": text})
        if "2级标题" in text:
            rules["text_rules"].append({"type": "heading_l2", "text": text})
        if "3级标题" in text:
            rules["text_rules"].append({"type": "heading_l3", "text": text})

        # Author rule
        if "作者" in text and "4号楷体" in text:
            rules["text_rules"].append({"type": "author", "text": text})

        # Punctuation rule
        if "标点" in text and ("中文" in text or "西文" in text):
            rules["text_rules"].append({"type": "punctuation", "text": text})

        # Table rule
        if "3线表" in text or "三线表" in text:
            rules["text_rules"].append({"type": "table_style", "text": text})

        # English author rule
        if "姓前名后" in text and "大写" in text and "英文" not in text:
            rules["text_rules"].append({"type": "author_format", "text": text})

        # Reference rules
        if "参考文献" in text and ("顺序" in text or "编码" in text):
            rules["text_rules"].append({"type": "reference_order", "text": text})
        if "中文文献" in text and "双语" in text:
            rules["text_rules"].append({"type": "bilingual_refs", "text": text})

        # English abstract rule
        if "Abstract" in text and ("第三人称" in text or "This paper" in text):
            rules["text_rules"].append({"type": "english_abstract", "text": text})

        # Introduction rule
        if "引言" in text and "不计入" in text:
            rules["text_rules"].append({"type": "introduction", "text": text})

        # Units and symbols
        if "GB 3100" in text or "GB3100" in text:
            rules["text_rules"].append({"type": "units_standard", "text": text})

    # ── Sample fonts from the template ──
    rules["font_samples"] = []
    seen = set()
    for key, info in font_samples.items():
        if key not in seen:
            seen.add(key)
            rules["font_samples"].append(info)

    return rules


def build_hardcoded_rules():
    """Return the hardcoded default rules for 深大学报理工版 template."""
    return {
        "page": {
            "width_cm": 21.0,
            "height_cm": 29.7,
            "top_margin_cm": 3.6,
            "bottom_margin_cm": 1.9,
            "left_margin_cm": 1.8,
            "right_margin_cm": 1.8,
        },
        "body_font": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 10.5},
        "heading_l1": {"chinese": "仿宋", "english": "Times New Roman", "size_pt": 16, "bold": False},
        "heading_l2": {"chinese": "黑体", "english": "Times New Roman", "size_pt": 10.5, "bold": True},
        "heading_l3": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5, "bold": False},
        "author": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 14},
        "abstract": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5},
        "keywords": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5},
        "title": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 22},
        "caption": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 9},
        "reference_title": {"chinese": "黑体", "english": "Times New Roman", "size_pt": 10.5},
        "reference_body": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 9},
        "line_spacing": 1.5,
        "first_line_indent_cm": 0.74,  # ≈2 Chinese characters at 5号
        "punctuation": {
            "chinese_full_stop": "。",   # U+3002 — standard Chinese period
            # Use "．" (U+FF0E) only if required by the template (e.g., 深大学报理工版)
            "chinese_comma": "，",
            "chinese_colon": "：",
            "chinese_semicolon": "；",
            "keyword_separator": "；",
            "auto_fix": True,    # Auto-convert English punct → Chinese in Chinese text
            "protect_numbers": True,  # Don't convert decimal points
            "protect_urls": True,     # Don't convert URLs
        },
        "table_style": "three_line",
        "reference_format": "sequential",  # [1], [2], ...
        "bilingual_required": True,
    }


def main():
    if len(sys.argv) < 2:
        print("Usage: extract_rules.py <template.docx> [output.json]")
        print("       extract_rules.py --default [output.json]  (use built-in default rules)")
        sys.exit(1)

    if sys.argv[1] == "--default":
        rules = build_hardcoded_rules()
        output_path = sys.argv[2] if len(sys.argv) > 2 else "rules.json"
    else:
        template_path = sys.argv[1]
        if not os.path.exists(template_path):
            print(f"Error: template not found: {template_path}")
            sys.exit(1)

        extracted = extract_rules(template_path)
        defaults = build_hardcoded_rules()
        # Merge: extracted provides page layout, defaults provide font mapping
        rules = {**defaults, "extracted": extracted}

        output_path = sys.argv[2] if len(sys.argv) > 2 else "rules.json"

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)

    print(f"Rules saved to: {output_path}")
    print(f"  Page: {rules.get('page', rules.get('extracted', {}).get('page', 'N/A'))}")
    print(f"  Body font: {rules.get('body_font', {})}")
    print(f"  Text rules found: {len(rules.get('text_rules', rules.get('extracted', {}).get('text_rules', [])))}")


if __name__ == '__main__':
    main()
