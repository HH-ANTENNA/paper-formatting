#!/usr/bin/env python3
"""
extract_rules.py — Extract formatting rules from a paper template .docx

Reads a Word template and produces a JSON rules file containing:
  1. Text rules (parsed from the template's instructional content)
  2. Visual formatting (observed font sizes, margins, etc.)

A built-in default (深大学报理工版 / 电子工艺报告 merged) is provided when
no custom template is given. Use --default to output the built-in rules directly.
"""
import json
import sys
import os
from docx import Document
from docx.shared import Cm, Pt, Emu


def emu_to_cm(emu):
    if emu is None:
        return None
    return round(emu / 360000, 2)


def emu_to_pt(emu):
    if emu is None:
        return None
    return round(emu / 12700, 1)


def extract_rules(template_path):
    """Extract all formatting rules from a custom template."""
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
        "text_rules": [],
    }

    section = doc.sections[0]
    rules["page"] = {
        "width_cm": emu_to_cm(section.page_width),
        "height_cm": emu_to_cm(section.page_height),
        "top_margin_cm": emu_to_cm(section.top_margin),
        "bottom_margin_cm": emu_to_cm(section.bottom_margin),
        "left_margin_cm": emu_to_cm(section.left_margin),
        "right_margin_cm": emu_to_cm(section.right_margin),
    }

    font_samples = {}

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        style_name = p.style.name
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

        # Parse text rules from instructional content
        if "正文叙述用" in text and "号" in text:
            rules["text_rules"].append({"type": "body_font", "text": text})
        if "中文题目" in text and "字" in text:
            rules["text_rules"].append({"type": "chinese_title", "text": text})
        if "英文标题" in text or "英文题目" in text:
            rules["text_rules"].append({"type": "english_title", "text": text})
        if "摘" in text and "要" in text and ("第三人称" in text or "5号楷体" in text or "撰写" in text):
            rules["text_rules"].append({"type": "abstract", "text": text})
        if "关键词" in text and "6" in text and "12" in text:
            rules["text_rules"].append({"type": "keywords", "text": text})
        if "1级标题" in text:
            rules["text_rules"].append({"type": "heading_l1", "text": text})
        if "2级标题" in text:
            rules["text_rules"].append({"type": "heading_l2", "text": text})
        if "3级标题" in text:
            rules["text_rules"].append({"type": "heading_l3", "text": text})
        if "作者" in text and "4号楷体" in text:
            rules["text_rules"].append({"type": "author", "text": text})
        if "标点" in text and ("中文" in text or "西文" in text):
            rules["text_rules"].append({"type": "punctuation", "text": text})
        if "3线表" in text or "三线表" in text:
            rules["text_rules"].append({"type": "table_style", "text": text})
        if "姓前名后" in text and "大写" in text and "英文" not in text:
            rules["text_rules"].append({"type": "author_format", "text": text})
        if "参考文献" in text and ("顺序" in text or "编码" in text):
            rules["text_rules"].append({"type": "reference_order", "text": text})
        if "中文文献" in text and "双语" in text:
            rules["text_rules"].append({"type": "bilingual_refs", "text": text})
        if "Abstract" in text and ("第三人称" in text or "This paper" in text):
            rules["text_rules"].append({"type": "english_abstract", "text": text})
        if "引言" in text and "不计入" in text:
            rules["text_rules"].append({"type": "introduction", "text": text})
        if "GB 3100" in text or "GB3100" in text:
            rules["text_rules"].append({"type": "units_standard", "text": text})

    rules["font_samples"] = []
    seen = set()
    for key, info in font_samples.items():
        if key not in seen:
            seen.add(key)
            rules["font_samples"].append(info)

    return rules


def build_default_rules():
    """Return the default formatting rules.

    Based on the 电子工艺报告 (Electronic Process Report) template as the
    authoritative source, with journal-paper features retained as options.
    This is the one and only default — no template switching needed.
    """
    return {
        "template_name": "深大学报理工版（默认）",
        # ── Page ──
        "page": {
            "width_cm": 21.0,
            "height_cm": 29.7,
            "top_margin_cm": 2.54,
            "bottom_margin_cm": 2.54,
            "left_margin_cm": 3.17,
            "right_margin_cm": 3.17,
        },

        # ── Body text: 宋体 小四 (11pt), 1.15x, justified ──
        "body_font": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 11},
        "line_spacing": 1.15,
        "first_line_indent_cm": 0.74,

        # ── Title: 宋体 小二 (18pt) Bold, Center ──
        "title": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 18, "bold": True},

        # ── Author / affiliation (用于期刊模式，可选) ──
        "author": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 14},
        "affiliation": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 9},

        # ── Abstract: 楷体 五号 (10.5pt), inline "摘要：" label ──
        "abstract_label": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 12},
        "abstract": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5},

        # ── Keywords: 楷体, inline "关键词：" Bold label, ；separator ──
        "keywords_label": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 12, "bold": True},
        "keywords": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5},

        # ── Headings ──
        "heading_l1": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 14, "bold": True},
        "heading_l2": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 15, "bold": True},
        "heading_l3": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 14, "bold": False},

        # ── Figure / Table caption: 宋体 小五 (9pt), Center ──
        "caption": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 9},

        # ── Reference title: 楷体 五号 (10.5pt) ──
        "reference_title": {"chinese": "楷体", "english": "Times New Roman", "size_pt": 10.5, "bold": False},
        # ── Reference body: 宋体 小五 (9pt), hanging indent ──
        "reference_body": {"chinese": "宋体", "english": "Times New Roman", "size_pt": 9},
        "reference_hanging_cm": 0.74,

        # ── Tables ──
        "table_style": "three_line",
        "reference_format": "sequential",

        # ── Bilingual (default: off — enable for journal submissions) ──
        "bilingual_required": False,

        # ── Section numbering (default: all sections numbered) ──
        "introduction_numbered": True,

        # ── Abstract format (【摘要】 bracketed inline label) ──
        "abstract_format": "bracket_label_inline",

        # ── Keywords ──
        "keyword_separator": "；",

        # ── Punctuation ──
        "punctuation": {
            "chinese_full_stop": "。",
            "chinese_comma": "，",
            "chinese_colon": "：",
            "chinese_semicolon": "；",
            "keyword_separator": "；",
            "auto_fix": True,
            "protect_numbers": True,
            "protect_urls": True,
        },

        # ── Reference indent ──
        "reference_indent": "hanging",
    }


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="Extract formatting rules from a paper template .docx"
    )
    parser.add_argument(
        "path",
        nargs="?",
        help="Template .docx path, OR output .json path when using --default",
    )
    parser.add_argument(
        "output",
        nargs="?",
        default="rules.json",
        help="Output JSON file path (default: rules.json; only used when extracting from template)",
    )
    parser.add_argument(
        "--default",
        action="store_true",
        help="Output the built-in default rules (no template file needed)",
    )
    args = parser.parse_args()

    if args.default:
        # --default mode: args.path (if given) is the output file
        output_path = args.path if args.path else "rules.json"
        rules = build_default_rules()
    elif args.path and os.path.exists(args.path):
        # Extract from a template file
        template_path = args.path
        extracted = extract_rules(template_path)
        defaults = build_default_rules()
        rules = {**defaults, "extracted": extracted}
        output_path = args.output
    elif not args.path:
        # No args at all → default rules
        rules = build_default_rules()
        output_path = "rules.json"
    else:
        print(f"Error: file not found: {args.path}")
        print("Use --default for built-in rules, or provide a valid template .docx")
        sys.exit(1)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)

    print(f"Rules saved to: {output_path}")
    print(f"  Template: {rules.get('template_name', 'Custom')}")
    print(f"  Page: {rules.get('page', {})}")
    print(f"  Body: {rules.get('body_font', {})}")
    print(f"  Bilingual: {rules.get('bilingual_required', False)}")
    extracted = rules.get("extracted", {})
    if extracted:
        print(f"  Text rules found: {len(extracted.get('text_rules', []))}")


if __name__ == '__main__':
    main()
