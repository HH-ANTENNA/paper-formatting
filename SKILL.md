---
name: paper-formatting
description: Use when the user needs to format an academic paper (.docx or .tex) according to a journal or university template — symptoms include "format my paper", "adjust my thesis", "排版", "论文格式", "apply template", or when the user has a paper draft that needs to conform to submission guidelines. Handles font sizes, margins, heading styles, three-line tables, bilingual elements, references, and punctuation.
---

# Paper Formatting Skill

## Overview

Format academic papers according to a unified default template (基于电子工艺报告 / 深大学报理工版合并)。The skill READS the template's written rules and visual formatting, then applies those rules to the user's draft. Supports Word (.docx) and LaTeX (.tex).

**One default template** — no switching needed. 以电子工艺报告的实测格式为准，期刊特性（双语摘要/关键词、作者拼音、中图分类号等）作为可选功能保留。

## When to Use

- User has a paper draft and needs it formatted for submission
- User mentions "格式", "排版", "模板", "format", "template"
- Symptoms: wrong fonts, wrong margins, tables not three-line, references wrong format, inconsistent punctuation
- Course reports, project reports, journal papers — all use the same formatting engine

## Core Workflow

```
User: "format my paper" or "帮我排版"
    │
    ▼
Step 1: CONFIRM INPUT
    "Which file needs formatting?"
    → User provides path to draft .docx
    │
    ▼
Step 2: GENERATE / LOAD RULES
    python scripts/extract_rules.py --default rules.json
    (Optionally: python scripts/extract_rules.py template.docx rules.json)
    → Show rule summary:
      "Page: A4, margins T/B/L/R = 2.54/2.54/3.17/3.17cm"
      "Body: 宋体/Times New Roman 11pt, 1.15x, indent 0.74cm"
    → "OK? (y/n)"
    │
    ▼
Step 3: ANALYZE TARGET DOCUMENT
    Quick scan of the user's draft
    → "Detected: 1 title, 1 abstract, 1 keywords,
      6 headings (1 L1, 3 L2, 2 L3), 10 figures, 2 refs"
    → "Proceed? (y/n)"
    │
    ▼
Step 4: FORMAT (phased, user confirms each phase)
    Phase A: Page setup + body font + paragraph + punctuation
    Phase B: Front matter (title, abstract, keywords)
    Phase C: Heading hierarchy (L1/L2/L3)
    Phase D: Tables, figures, equations
    Phase E: References
    │
    ▼
Step 5: VERIFY
    python scripts/verify_format.py <output.docx> rules.json
    → Report: ✅ Fixed + ⚠️ Warnings + ❌ Manual fixes
```

---

## Default Template Rules

Based on the 电子工艺报告 (Electronic Process Report) format, with journal-paper support retained.

### Page & Body
- Paper: A4 (21cm × 29.7cm)
- Margins: top=2.54cm, bottom=2.54cm, left=3.17cm, right=3.17cm
- Body text: 宋体 (SimSun), 小四 (11pt), 1.15x line spacing, justified
- First-line indent: 2 characters (≈0.74cm at 小四)

### Font Hierarchy
| Element | Font | Size | Bold | Align |
|---------|------|------|------|-------|
| Title | 宋体 (SimSun) | 小二 (18pt) | Bold | Center |
| Abstract label "【摘要】" | 楷体 (KaiTi) | 小四 (12pt) | — | Left |
| Abstract body | 楷体 | 五号 (10.5pt) | — | Justified |
| Keywords label "【关键词】" | 楷体 | 小四 (12pt) | Bold | Left |
| Keywords items | 楷体 | 五号 (10.5pt) | — | Left |
| L1 heading "1  XXX" | 宋体 | 四号 (14pt) | Bold | Left |
| L2 heading "1.1  XXX" | 宋体 | 小三 (15pt) | Bold | Left |
| L3 heading "1.1.1  XXX" | 楷体 | 四号 (14pt) | — | Left |
| Figure caption "图X  ..." | 宋体 | 小五 (9pt) | — | Center |
| Table caption "表X  ..." | 宋体 | 小五 (9pt) | — | Center |
| Reference title "参考文献" | 楷体 | 五号 (10.5pt) | — | Left |
| Reference body [1]... | 宋体 | 小五 (9pt) | — | Justified |

### Headings
- ALL sections are numbered: including 引言 (1), 参考文献 (last number)
- L1: `"1  引言"` (number + 2 spaces + title)
- L2: `"1.1  项目背景"` (number + 2 spaces + title)
- L3: `"1.1.1  经济可行性"` (number + 1~2 spaces + title)
- Use Heading 1 / Heading 2 / Heading 3 styles

### Abstract & Keywords
- Abstract: `【摘要】` + content (same paragraph, inline format, black lenticular brackets U+3010/U+3011)
  - `【摘要】` label in 楷体 12pt; body in 楷体 10.5pt
  - Also accepts `摘要：` format and normalizes to `【摘要】`
- Keywords: `【关键词】` + items separated by `；` (fullwidth semicolon)
  - `【关键词】` label in 楷体 12pt Bold; items in 楷体 10.5pt
  - Also accepts `关键词：` format and normalizes to `【关键词】`

### Tables
- Three-line table style (top thick 1.5pt, header-bottom thin 0.75pt, bottom thick 1.5pt)
- No vertical borders, no interior horizontal borders
- Caption: "表X  标题" (宋体 小五 9pt), centered, above the table
- Units in header: "量符号/单位" format

### Figures
- Caption: "图X  标题" (宋体 小五 9pt), centered, below the figure
- Max 6 figures recommended

### References
- Title: "参考文献" (楷体 10.5pt)
- Sequential numbering: [1], [2], [3]...
- GB/T 7714 format
- Hanging indent: first line -0.74cm, left indent 0.74cm
- Body: 宋体 小五 (9pt)

### Punctuation
- **Context-aware**: Chinese text → Chinese punctuation; English text → English punctuation
  - `.` after CJK char → `。`  |  `.` after Latin letter → `.` (kept)
  - `,` → `，`  |  `:` → `：`  |  `;` → `；`  |  `?` → `？`  |  `!` → `！`
  - `(` → `（`  |  `)` → `）`
- **Protected**: URLs, email, decimal numbers (3.14), abbreviations (i.e., etc.), numbered lists
- Full stop: `"。"` (U+3002)

### Bilingual (Optional — enable for journal submissions)
- Title: Chinese + English
- Authors: Chinese + English (pinyin)
- Abstract: Chinese + English
- Keywords: Chinese + English
- Table/figure captions: Chinese + English
- Chinese refs: Chinese + English with "(in Chinese)"
- CLC number / Document Code for journal papers
- Set `"bilingual_required": true` in rules.json to enable

---

## Scripts

### extract_rules.py
Outputs the default rules JSON, or extracts from a custom template.

```bash
# Default rules (recommended):
python scripts/extract_rules.py --default rules.json

# Extract from custom template:
python scripts/extract_rules.py template.docx rules.json
```

### format_docx.py
Applies formatting rules to a .docx file in phases.

```bash
python scripts/format_docx.py input.docx rules.json -o output.docx
python scripts/format_docx.py input.docx rules.json -o output.docx --phase B  # front matter only
python scripts/format_docx.py input.docx rules.json -o output.docx --phase A,C  # page + headings only
```

### format_latex.py
Applies formatting rules to a .tex file.

```bash
python scripts/format_latex.py input.tex rules.json -o output.tex
```

### verify_format.py
Checks formatted document against rules.

```bash
python scripts/verify_format.py output.docx rules.json
python scripts/verify_format.py output.docx rules.json --json  # machine-readable
```

---

## Implementation Notes

### Setting CJK + Latin fonts on a run
```python
def set_run_font(run, chinese_font, english_font, size_pt, bold=None):
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
```

### Three-line table
- Remove all table and cell borders → apply only top (1.5pt) + bottom (0.75pt) on header row, bottom (1.5pt) on last row
- All other borders set to `nil`

### Inline abstract detection
```python
if re.match(r'^摘\s*要[：:]', text):
    # Split into label ("摘要：") and body — different fonts
```

### Heading styles & collapse/expand
Phase C applies Word built-in **Heading 1/2/3** styles (not just font formatting). This gives headings the native collapse/expand triangle in Word's navigation pane — same behavior as the template.
```python
# Apply Heading 1 style (enables collapse/expand), then override font
p.style = doc.styles['Heading 1']
for run in p.runs:
    set_run_font(run, "宋体", "Times New Roman", 14, bold=True)
```

---

## Common Mistakes to Avoid

1. **Don't just copy visual font** — read the text rules and understand intent
2. **Don't convert all periods blindly** — periods after Latin letters are kept; only after CJK chars → `。`
3. **Don't leave tables as Table Grid** — convert to three-line
4. **Don't use Chinese punctuation in English text** — and vice versa
5. **All sections are numbered** — including 引言 and 参考文献
6. **Abstract uses 【】 brackets** — `【摘要】` label in same paragraph as body text
7. **Keywords use `；`** — not commas
8. **References use hanging indent** — first line back 0.74cm
9. **Don't format in one shot** — phase the work, confirm each phase
10. **Bilingual is OFF by default** — enable only for journal submissions

## Red Flags — STOP and Ask User

- Template file cannot be read → verify path
- Document structure ambiguous → ask user to clarify
- >50% rules cannot be auto-applied → present manual checklist
- python-docx not installed → `pip install python-docx`
