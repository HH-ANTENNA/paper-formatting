#!/usr/bin/env python3
"""
verify_format.py — Check a formatted document against template rules.

Reports issues in two categories:
  ✅ Fixed: issues that were automatically corrected
  ⚠️ Manual: issues requiring human attention
"""
import json
import sys
import re
from docx import Document
from docx.shared import Cm, Pt


def check_document(doc_path, rules):
    """Check document against rules, return report."""
    doc = Document(doc_path)
    report = {"passed": [], "warnings": [], "manual_fixes": []}

    body = rules.get("body_font", {})
    l1 = rules.get("heading_l1", {})
    l2 = rules.get("heading_l2", {})
    l3 = rules.get("heading_l3", {})

    # Check page size
    for i, section in enumerate(doc.sections):
        w_cm = section.page_width / 360000
        h_cm = section.page_height / 360000
        if abs(w_cm - 21.0) > 0.5 or abs(h_cm - 29.7) > 0.5:
            report["warnings"].append(f"Section {i}: page size {w_cm:.1f}x{h_cm:.1f}cm (expected A4)")

    # Check paragraphs
    has_english_abstract = False
    has_english_keywords = False
    has_chinese_refs = False
    has_english_refs = False

    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue

        # Check for bilingual elements
        if text.lower().startswith('abstract'):
            has_english_abstract = True
        if text.lower().startswith('key words') or text.lower().startswith('keywords'):
            has_english_keywords = True

        # Check reference format
        if re.match(r'^\[\d+\]', text):
            # Chinese characters in ref suggest bilingual needed
            if any('一' <= c <= '鿿' for c in text):
                has_chinese_refs = True
            else:
                has_english_refs = True

        # Check for "。" usage
        if '。' in text:
            report["warnings"].append(
                f"P[{i}]: Contains '。' (U+3002) — should use '．' (U+FF0E): "
                f"{text[:60]}..."
            )

        # Check for English punctuation in Chinese text
        if any('一' <= c <= '鿿' for c in text):
            if re.search(r'(?<=[一-鿿])[,]', text):
                report["warnings"].append(f"P[{i}]: Chinese text uses English comma: {text[:60]}...")

    # Bilingual checks
    if not has_english_abstract and rules.get("bilingual_required", True):
        report["manual_fixes"].append("Missing English abstract — template requires bilingual abstract")
    if not has_english_keywords and rules.get("bilingual_required", True):
        report["manual_fixes"].append("Missing English keywords — template requires bilingual keywords")

    # Table checks
    for ti, table in enumerate(doc.tables):
        # Check for three-line style
        tblPr = table._tbl.tblPr
        if tblPr is not None:
            borders = tblPr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders')
            # Simplified check
        report["manual_fixes"].append(
            f"Table {ti}: Verify three-line style applied correctly. "
            f"Check bilingual caption exists above table."
        )

    # Reference checks
    if has_chinese_refs and not has_english_refs:
        report["manual_fixes"].append(
            "Chinese references detected without English translations — "
            "template requires bilingual (Chinese + English) for all Chinese references"
        )

    return report


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Verify formatted academic paper")
    parser.add_argument("document", help="Formatted .docx file")
    parser.add_argument("rules", help="Rules JSON file")
    parser.add_argument("--json", action="store_true", help="Output as JSON")
    args = parser.parse_args()

    with open(args.rules, 'r', encoding='utf-8') as f:
        rules = json.load(f)

    report = check_document(args.document, rules)

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print("=" * 60)
        print("FORMAT VERIFICATION REPORT")
        print("=" * 60)

        print(f"\n[PASSED] ({len(report['passed'])})")
        for item in report["passed"]:
            print(f"  [OK] {item}")
        if not report["passed"]:
            print("  (no automated checks passed)")

        print(f"\n[WARNINGS] ({len(report['warnings'])})")
        for item in report["warnings"]:
            print(f"  [WARN] {item}")
        if not report["warnings"]:
            print("  None")

        print(f"\n[MANUAL FIXES NEEDED] ({len(report['manual_fixes'])})")
        for item in report["manual_fixes"]:
            print(f"  [TODO] {item}")
        if not report["manual_fixes"]:
            print("  None")

        print("\n" + "=" * 60)


if __name__ == '__main__':
    main()
