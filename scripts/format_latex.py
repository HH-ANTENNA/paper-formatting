#!/usr/bin/env python3
"""
format_latex.py — Format a LaTeX paper according to the default template rules.

Modifies the preamble and applies formatting conventions.
"""
import json
import sys
import os
import re


def apply_latex_formatting(content, rules):
    changes = []
    bilingual = rules.get("bilingual_required", False)

    # ── Page geometry ──
    page = rules.get("extracted", {}).get("page", rules.get("page", {}))
    if page:
        geo_cmd = (
            f"\\geometry{{"
            f"a4paper,"
            f"top={page.get('top_margin_cm', 2.54)}cm,"
            f"bottom={page.get('bottom_margin_cm', 2.54)}cm,"
            f"left={page.get('left_margin_cm', 3.17)}cm,"
            f"right={page.get('right_margin_cm', 3.17)}cm"
            f"}}"
        )
        if '\\geometry{' not in content:
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
        font_settings.append(
            f"\\setmainfont{{{body.get('english', 'Times New Roman')}}}"
        )
        font_settings.append(
            f"\\setCJKmainfont{{{body.get('chinese', 'SimSun')}}}"
        )

    # Heading fonts
    l1 = rules.get("heading_l1", {})
    l2 = rules.get("heading_l2", {})
    l3 = rules.get("heading_l3", {})

    if l1:
        font_settings.append(
            f"\\CTEXsetup[nameformat={{\\zihao{{4}}\\bfseries}}]{{section}}"
        )
    if l2:
        font_settings.append(
            f"\\CTEXsetup[nameformat={{\\zihao{{3}}\\bfseries}}]{{subsection}}"
        )
    if l3:
        font_settings.append(
            f"\\CTEXsetup[nameformat={{\\zihao{{4}}\\kaishu}}]{{subsubsection}}"
        )

    # Line spacing
    ls = rules.get("line_spacing", 1.15)
    font_settings.append(f"\\linespread{{{ls}}}")

    # Paragraph indent
    indent = rules.get("first_line_indent_cm", 0.74)
    font_settings.append(f"\\setlength{{\\parindent}}{{{indent}cm}}")

    # Abstract setup
    font_settings.append("\\ctexset{abstractname={摘要}}")

    # Inject
    for fs in reversed(font_settings):
        if fs not in content:
            if '\\geometry{' in content:
                content = content.replace('\\geometry{', f'{fs}\n\\geometry{{')
            else:
                content = re.sub(
                    r'(\\documentclass(?:\[.*?\])?\{.*?\})',
                    f'\\1\n{fs}',
                    content
                )
            changes.append(f"Added: {fs}")

    # ── Punctuation ──
    punctuation_rules = rules.get("punctuation", {})
    full_stop = punctuation_rules.get("chinese_full_stop", "。")

    protected_patterns = [
        (r'https?://[a-zA-Z0-9._~:/?#\[\]@!$&()*+;=%\-]+', 'URL'),
        (r'\\[a-zA-Z]+\{[^}]*\}', 'LATEX_CMD'),
        (r'\\[a-zA-Z]+', 'LATEX_MACRO'),
        (r'\$\$?[^$]+\$\$?', 'MATH'),
        (r'\(\d+(?:\.\d+)?\)', 'PAREN_NUM'),
        (r'\d+\.\d+', 'DECIMAL'),
    ]
    placeholders = {}
    counter = [0]
    for pattern, ptype in protected_patterns:
        def make_replacer():
            def replacer(m):
                ph = f'\x00PROTECT{ptype}\x00{counter[0]}\x00'
                placeholders[ph] = m.group(0)
                counter[0] += 1
                return ph
            return replacer
        content = re.sub(pattern, make_replacer(), content)

    punct_map = {
        ',': '，', ':': '：', ';': '；', '?': '？', '!': '！',
        '(': '（', ')': '）',
    }
    for en_punct, cn_punct in punct_map.items():
        escaped = re.escape(en_punct)
        content = re.sub(f'(?<=[一-鿿]){escaped}', cn_punct, content)
        content = re.sub(f'{escaped}(?=[一-鿿])', cn_punct, content)

    content = re.sub(r'(?<=[一-鿿])\.(?=\s|$|[一-鿿])', full_stop, content)

    for ph, original in placeholders.items():
        content = content.replace(ph, original)

    # ── References ──
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

    print(f"LaTeX formatted. {len(changes)} changes:")
    for c in changes:
        print(f"  ✓ {c}")
    print(f"Saved: {args.output}")


if __name__ == '__main__':
    main()
