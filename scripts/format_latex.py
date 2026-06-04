#!/usr/bin/env python3
"""
format_latex.py — Format a LaTeX paper according to template rules.

Modifies the preamble and applies formatting conventions.
"""
import json
import sys
import os
import re


def apply_latex_formatting(content, rules):
    """Apply formatting rules to LaTeX source."""
    changes = []
    lines = content.split('\n')

    # ── Page geometry ──
    page = rules.get("extracted", {}).get("page", rules.get("page", {}))
    if page:
        geo_cmd = (
            f"\\geometry{{"
            f"a4paper,"
            f"top={page.get('top_margin_cm', 3.6)}cm,"
            f"bottom={page.get('bottom_margin_cm', 1.9)}cm,"
            f"left={page.get('left_margin_cm', 1.8)}cm,"
            f"right={page.get('right_margin_cm', 1.8)}cm"
            f"}}"
        )
        # Check if geometry is already in preamble
        if '\\geometry{' not in content:
            # Insert after documentclass
            content = re.sub(
                r'(\\documentclass(?:\[.*?\])?\{.*?\})',
                f'\\1\n{geo_cmd}',
                content
            )
            changes.append(f"Added geometry: {geo_cmd}")

    # ── Font settings ──
    body = rules.get("body_font", {})
    font_settings = []
    if body:
        font_settings.append(f"\\setmainfont{{{body.get('english', 'Times New Roman')}}}")
        font_settings.append(f"\\setCJKmainfont{{{body.get('chinese', 'SimSun')}}}")

    # Heading fonts
    l1 = rules.get("heading_l1", {})
    l2 = rules.get("heading_l2", {})
    l3 = rules.get("heading_l3", {})

    if l1 or l2 or l3:
        font_settings.append(
            f"\\CTEXsetup[nameformat={{\\zihao{{3}}\\fangsong}}]{{section}}"
        )
        font_settings.append(
            f"\\CTEXsetup[nameformat={{\\zihao{{5}}\\heiti}}]{{subsection}}"
        )
        font_settings.append(
            f"\\CTEXsetup[nameformat={{\\zihao{{5}}\\kaishu}}]{{subsubsection}}"
        )

    # ── Line spacing ──
    ls = rules.get("line_spacing", 1.5)
    font_settings.append(f"\\linespread{{{ls}}}")

    # ── Paragraph indent ──
    indent = rules.get("first_line_indent_cm", 0.74)
    # 2 characters ≈ 0.74cm at 5号
    font_settings.append(f"\\setlength{{\\parindent}}{{{indent}cm}}")

    # Inject font settings after geometry
    for fs in reversed(font_settings):
        if fs not in content:
            # Add after geometry or documentclass
            if '\\geometry{' in content:
                content = content.replace('\\geometry{', f'{fs}\n\\geometry{{')
            else:
                content = re.sub(
                    r'(\\documentclass(?:\[.*?\])?\{.*?\})',
                    f'\\1\n{fs}',
                    content
                )
            changes.append(f"Added: {fs}")

    # ── Punctuation fixes ──
    # Replace "。" with "．" in Chinese text blocks
    content = content.replace('。', '．')

    # ── Reference format ──
    # Ensure biblatex/natbib settings for sequential numbering
    if '\\usepackage' in content and 'biblatex' not in content and 'natbib' not in content:
        ref_pkg = '\\usepackage[sort&compress,numbers]{natbib}'
        content = content.replace(
            '\\begin{document}',
            f'{ref_pkg}\n\\begin{{document}}'
        )
        changes.append("Added natbib for reference formatting")

    return content, changes


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Format academic paper (.tex)")
    parser.add_argument("input", help="Input .tex file")
    parser.add_argument("rules", help="Rules JSON file")
    parser.add_argument("--output", "-o", required=True, help="Output .tex file")
    args = parser.parse_args()

    with open(args.rules, 'r', encoding='utf-8') as f:
        rules = json.load(f)

    with open(args.input, 'r', encoding='utf-8') as f:
        content = f.read()

    new_content, changes = apply_latex_formatting(content, rules)

    with open(args.output, 'w', encoding='utf-8') as f:
        f.write(new_content)

    print(f"LaTeX formatting complete. {len(changes)} changes made:")
    for c in changes:
        print(f"  ✓ {c}")
    print(f"Saved to: {args.output}")


if __name__ == '__main__':
    main()
