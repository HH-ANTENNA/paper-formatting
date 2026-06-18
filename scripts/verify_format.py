#!/usr/bin/env python3
"""
verify_format.py — Check a formatted document against template rules.

Reports issues in three categories:
  ✅ Passed: rules that are correctly applied
  ⚠️ Warnings: possible issues that need review
  ❌ Manual: issues requiring human attention
"""
import json
import sys
import re
from docx import Document
from docx.shared import Cm, Pt


def is_chinese_char(ch):
    """Check if a character is in the CJK range."""
    cp = ord(ch)
    return (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or
            0x20000 <= cp <= 0x2A6DF or 0xF900 <= cp <= 0xFAFF or
            0x3000 <= cp <= 0x303F or 0xFF00 <= cp <= 0xFFEF)


# English punctuation that should NOT appear in Chinese-dominant text
_EN_PUNCT_IN_CN_WARNING = {
    ',': ('，(U+FF0C)', 'English comma "," in Chinese text'),
    ';': ('；(U+FF1B)', 'English semicolon ";" in Chinese text'),
    ':': ('：(U+FF1A)', 'English colon ":" in Chinese text after Chinese char'),
    '?': ('？(U+FF1F)', 'English question mark "?" in Chinese text'),
    '!': ('！(U+FF01)', 'English exclamation mark "!" in Chinese text'),
    '(': ('（(U+FF08)', 'English left parenthesis "(" in Chinese text'),
    ')': ('）(U+FF09)', 'English right parenthesis ")" in Chinese text'),
}


def check_document(doc_path, rules):
    """Check document against rules, return report."""
    doc = Document(doc_path)
    report = {"passed": [], "warnings": [], "manual_fixes": []}

    body = rules.get("body_font", {})
    l1 = rules.get("heading_l1", {})
    l2 = rules.get("heading_l2", {})
    l3 = rules.get("heading_l3", {})

    # Full stop configuration
    punctuation_rules = rules.get("punctuation", {})
    expected_full_stop = punctuation_rules.get("chinese_full_stop", "。")
    full_stop_name = "．(U+FF0E)" if expected_full_stop == "．" else "。(U+3002)"

    # Check page size
    for i, section in enumerate(doc.sections):
        w_cm = section.page_width / 360000
        h_cm = section.page_height / 360000
        if abs(w_cm - 21.0) > 0.5 or abs(h_cm - 29.7) > 0.5:
            report["warnings"].append(f"Section {i}: page size {w_cm:.1f}x{h_cm:.1f}cm (expected A4)")
        else:
            report["passed"].append(f"Section {i}: page size A4 ({w_cm:.1f}x{h_cm:.1f}cm)")

    # Check paragraphs
    has_english_abstract = False
    has_english_keywords = False
    has_chinese_refs = False
    has_english_refs = False
    punctuation_issues = []
    full_stop_issues = []

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

        # ── Punctuation checks for Chinese-dominant paragraphs ──
        # Determine if this paragraph is Chinese-dominant
        total_alphabetic = sum(1 for c in text if c.isalpha())
        chinese_chars = sum(1 for c in text if '一' <= c <= '鿿')
        is_chinese_dominant = chinese_chars > total_alphabetic * 0.5 and chinese_chars > 5

        if is_chinese_dominant:
            # Check for English punctuation that should be Chinese
            for en_punct, (cn_name, desc) in _EN_PUNCT_IN_CN_WARNING.items():
                if en_punct in text:
                    # For period, check context more carefully
                    if en_punct == '.':
                        continue  # Period handled separately below
                    # Find the punctuation and check context
                    for m in re.finditer(re.escape(en_punct), text):
                        pos = m.start()
                        prev_char = text[pos - 1] if pos > 0 else ''
                        next_char = text[pos + 1] if pos + 1 < len(text) else ''
                        prev_is_cjk = is_chinese_char(prev_char)
                        next_is_cjk = is_chinese_char(next_char)
                        if prev_is_cjk or next_is_cjk:
                            punctuation_issues.append(
                                f"P[{i}]: {desc}: "
                                f"\"...{text[max(0,pos-5):pos+6]}...\" → should use {cn_name}"
                            )
                            break  # One mention per paragraph per punct type

            # Check full stop: warn about "。" if template expects "．"
            if '。' in text and expected_full_stop == '．':
                # Count occurrences
                count = text.count('。')
                full_stop_issues.append(
                    f"P[{i}]: {count}x '。'(U+3002) found — template expects {full_stop_name}: "
                    f"\"{text[:60]}{'...' if len(text) > 60 else ''}\""
                )

            # Check full stop: warn about "．" if template expects "。"
            if '．' in text and expected_full_stop == '。':
                count = text.count('．')
                full_stop_issues.append(
                    f"P[{i}]: {count}x '．'(U+FF0E) found — template expects {full_stop_name}: "
                    f"\"{text[:60]}{'...' if len(text) > 60 else ''}\""
                )

        # Check for isolated English period in Chinese-dominant text
        # ('.' followed by space or newline, preceded by text)
        if is_chinese_dominant:
            for m in re.finditer(r'(?<=[^\d])\.(?=\s|$)', text):
                pos = m.start()
                prev_char = text[pos - 1] if pos > 0 else ''
                if is_chinese_char(prev_char):
                    punctuation_issues.append(
                        f"P[{i}]: English period '.' after Chinese char — should use "
                        f"{full_stop_name}: \"...{text[max(0,pos-5):pos+6]}...\""
                    )
                    break

    # ── Aggregate punctuation warnings ──
    if punctuation_issues:
        report["warnings"].extend(punctuation_issues)
    else:
        report["passed"].append("No English punctuation found in Chinese-dominant text")

    if full_stop_issues:
        report["manual_fixes"].extend(full_stop_issues)
    elif is_chinese_dominant:  # Only if we actually checked
        report["passed"].append(f"Full stop character matches template ({full_stop_name})")

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
