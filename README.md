# Paper Formatting Skill

A Claude Code skill for formatting academic papers according to journal/university templates. Supports both Word (.docx) and LaTeX (.tex).

## What It Does

Reads a paper-writing template (style guide), extracts the formatting rules from its content, and applies them to your draft — covering fonts, margins, heading hierarchy, three-line tables, bilingual captions, reference formatting, and punctuation.

## Quick Start

1. Copy this entire folder to `~/.claude/skills/paper-formatting/`
2. In Claude Code, say: **"Format my paper"** or **"帮我排版论文"**
3. The skill will ask:
   - Use default template? (y/n)
   - If no → provide your template path
4. Formatting runs in 5 phases with checkpoints between each:
   - **Phase A:** Page setup + body font + punctuation
   - **Phase B:** Front matter (title, authors, abstract, keywords)
   - **Phase C:** Heading hierarchy (L1/L2/L3)
   - **Phase D:** Tables (three-line) + figure captions
   - **Phase E:** References

## Default Template

The default template follows **深圳大学学报理工版 (Journal of Shenzhen University Science and Engineering)** submission guidelines:

| Element | Font | Size |
|---------|------|------|
| Body text | 宋体 / Times New Roman | 5号 (10.5pt) |
| L1 heading | 仿宋 | 3号 (16pt) |
| L2 heading | 黑体 | 5号 (10.5pt) |
| L3 heading | 楷体 | 5号 (10.5pt) |
| Authors | 楷体 | 4号 (14pt) |
| Abstract | 楷体 | 5号 (10.5pt) |
| Captions | 宋体 | 小五 (9pt) |

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/extract_rules.py` | Extract formatting rules from a template .docx → JSON |
| `scripts/format_docx.py` | Apply rules to Word documents (phased) |
| `scripts/format_latex.py` | Apply rules to LaTeX documents |
| `scripts/verify_format.py` | Check formatted document against rules |

### CLI Usage

```bash
# Extract rules from a template
python scripts/extract_rules.py template.docx rules.json

# Use built-in default rules
python scripts/extract_rules.py --default rules.json

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

## Custom Templates

To use your own template:
1. When the skill asks "Use default template?", answer **no**
2. Provide the path to your .docx template
3. The skill will extract rules from both the template's text instructions and its visual formatting

## License

MIT
