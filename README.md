# Paper Formatting Skill

[🇨🇳 中文文档](README_zh.md)

A Claude Code skill for formatting academic papers according to university templates. Supports Word (.docx) and LaTeX (.tex).

## ⚠️ Disclaimer / 免责声明

**This skill is for learning and reference purposes only.** While it automates common formatting tasks, it cannot cover every edge case. Formatting requirements vary between institutions, journals, and even individual reviewers. **Always manually review your formatted document** — check font sizes, heading levels, table styles, reference formatting, and page layout against the official guidelines. The authors assume no liability for formatting errors in submissions.

## What It Does

Reads a paper-writing template (style guide), extracts the formatting rules, and applies them to your draft — covering fonts, margins, heading hierarchy, three-line tables, figure captions, reference formatting, and punctuation.

## Quick Start

1. Copy this entire folder to `~/.claude/skills/paper-formatting/`
2. In Claude Code, say: **"Format my paper"** or **"帮我排版论文"**
3. The skill will apply formatting in 5 phases:
   - **Phase A:** Page setup + body font + punctuation
   - **Phase B:** Front matter (title, abstract 【摘要】, keywords 【关键词】)
   - **Phase C:** Heading hierarchy (L1/L2/L3 with Heading styles for collapse/expand)
   - **Phase D:** Tables (three-line) + figure captions
   - **Phase E:** References

## Default Template Specs

| Element | Font | Size | Spacing |
|---------|------|------|---------|
| Title | 宋体 (SimSun) | 小二 (18pt) Bold | Center |
| Body text | 宋体 / Times New Roman | 五号 (10.5pt) | Single, 段前段后0.5行 |
| Abstract 【摘要】 | 楷体 | label 12pt / body 10.5pt | — |
| Keywords 【关键词】 | 楷体 | label 12pt Bold / items 10.5pt | — |
| L1 heading | 黑体 (SimHei) Bold | 三号 (16pt) | Single, 段前段后0.5行 |
| L2 heading | 黑体 (SimHei) Bold | 小三 (15pt) | Single, 段前段后0.5行 |
| L3 heading | 黑体 (SimHei) Bold | 四号 (14pt) | Single, 段前段后0.5行 |
| Figure caption | 宋体 | 小五 (9pt) | Center |
| Table caption | 宋体 | 小五 (9pt) | Center |
| Reference title | 楷体 | 五号 (10.5pt) | Left |
| Reference body | 宋体 | 小五 (9pt) | Hanging indent |

- **Page:** A4, margins T/B/L/R = 2.54/2.54/3.17/3.17 cm
- **All headings** use Word built-in Heading 1/2/3 styles (enables collapse/expand in navigation pane)
- **All sections numbered:** including 引言 (e.g., "1  引言")
- **Abstract/Keywords:** 【】black lenticular brackets, inline format
- **Keywords separator:** fullwidth semicolon `；`

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/extract_rules.py` | Generate default rules JSON or extract from a template .docx |
| `scripts/format_docx.py` | Apply rules to Word documents (phased) |
| `scripts/format_latex.py` | Apply rules to LaTeX documents |
| `scripts/verify_format.py` | Check formatted document against rules |

### CLI Usage

```bash
# Generate default rules
python scripts/extract_rules.py --default rules.json

# Extract rules from a custom template
python scripts/extract_rules.py template.docx rules.json

# Format a document (all phases)
python scripts/format_docx.py paper.docx rules.json -o formatted.docx --phase all

# Format specific phases only
python scripts/format_docx.py paper.docx rules.json -o formatted.docx --phase A,B,C

# Verify formatting
python scripts/verify_format.py formatted.docx rules.json
```

## Requirements

```bash
pip install python-docx
```

## License

MIT
