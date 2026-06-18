---
name: paper-formatting
description: Use when the user needs to format an academic paper (.docx or .tex) according to a journal or university template — symptoms include "format my paper", "adjust my thesis", "排版", "论文格式", "apply template", or when the user has a paper draft that needs to conform to submission guidelines. Handles font sizes, margins, heading styles, three-line tables, bilingual elements, references, and punctuation.
---

# Paper Formatting Skill

## Overview

Format academic papers according to a style-guide template. The skill READS the template's written rules (not just copies visual formatting), then applies those rules to the user's draft. Supports Word (.docx) and LaTeX (.tex).

## When to Use

- User has a paper draft and needs it formatted for submission
- User mentions "格式", "排版", "模板", "format", "template"
- Symptoms: wrong fonts, wrong margins, tables not three-line, references wrong format, missing bilingual sections, inconsistent punctuation

## Core Workflow

```
User: "format my paper" or "帮我排版"
    │
    ▼
Step 1: CONFIRM TEMPLATE
    "Use default template (深大学报理工版)?"
    Yes → use 论文写作模板.docx
    No  → ask user to provide template path
    │
    ▼
Step 2: EXTRACT RULES from template
    Run extract_rules.py on the template
    → Show rule summary to user
    → "Does this look correct? (y/n)"
    │
    ▼
Step 3: ANALYZE TARGET DOCUMENT
    Run analyze step on the user's draft
    → Show: "Detected: 1 title, 3 authors, 1 abstract,
      5 headings (1 L1, 2 L2, 2 L3), 1 table, 3 refs"
    → "Proceed with formatting? (y/n)"
    │
    ▼
Step 4: FORMAT (phased, user confirms each phase)
    Phase A: Page setup + body font + paragraph + punctuation
    Phase B: Front matter (title, authors, abstract, keywords, bilingual)
    Phase C: Heading hierarchy (L1/L2/L3)
    Phase D: Tables, figures, equations
    Phase E: References
    │
    ▼
Step 5: VERIFY
    Run verify_format.py
    → Report: ✅ Fixed items + ⚠️ Manual fixes needed
```

## Template Rules (Default: 深大学报理工版)

### Page & Body
- Paper: A4 (21cm × 29.7cm)
- Margins: top=3.6cm, bottom=1.9cm, left=1.8cm, right=1.8cm (extract from template)
- Body text: 宋体 (SimSun), 5号 (10.5pt), 1.5x line spacing
- First-line indent: 2 characters

### Font Hierarchy
| Element | Font | Size |
|---------|------|------|
| Chinese title | 宋体 (SimSun) | 二号 (22pt) |
| Authors | 楷体 (KaiTi) | 四号 (14pt) |
| Affiliation | 宋体 | 小五 (9pt) |
| Abstract | 楷体 | 五号 (10.5pt) |
| Keywords | 楷体 | 五号 (10.5pt) |
| L1 heading | 仿宋 (FangSong) | 三号 (16pt) |
| L2 heading | 黑体 (SimHei) | 五号 (10.5pt) |
| L3 heading | 楷体 (KaiTi) | 五号 (10.5pt) |
| Table/figure caption | 宋体 | 小五 (9pt) |
| Reference body | 宋体 | 小五 (9pt) |
| Reference title | 黑体 | 五号 (10.5pt) |

### Headings
- Introduction (引言) has NO section number
- L1: "1 XXXX" format, 仿宋 三号
- L2: "1.1 XXXX" format, 黑体 五号
- L3: "1.1.1 XXXX" format, 楷体 五号
- List ordering: 1) 2) 3) first level, then ①②③

### Punctuation
- **Context-aware auto-detection**: Chinese text gets Chinese punctuation, English text keeps English punctuation
  - `.` after Chinese char → `。`  |  `.` after English letter → `.` (kept)
  - `,` → `，`  |  `:` → `：`  |  `;` → `；`  |  `?` → `？`  |  `!` → `！`
  - `(` → `（`  |  `)` → `）`
- **Protected from conversion**: URLs, email addresses, decimal numbers (3.14), abbreviations (i.e., etc.), numbered lists (1. 2. 3.)
- **Full stop character** (configurable in rules.json → `punctuation.chinese_full_stop`):
  - `"。"` (U+3002) — **Default**, standard Chinese period
  - `"．"` (U+FF0E) — Only if template requires it (e.g., 深大学报理工版)
- Keywords separated by `；` (fullwidth semicolon)

### Tables
- Three-line table style (top thick, header-bottom thin, bottom thick)
- No vertical borders, no interior horizontal borders
- Bilingual caption (Chinese above, English below)
- Caption font: 小五 (9pt) bold for "表X", normal for title text
- Units in header: "量符号/单位" format

### Figures
- Bilingual caption: "图X 中文标题" + "Fig. X English title"
- Max 6 figures recommended
- Figure caption below the figure

### References
- Sequential numbering: [1], [2], [3]...
- Chinese references MUST have English translation appended with "(in Chinese)"
- Author format: Surname ALL CAPS, given name abbreviated (no dots)
- 3+ authors: use "等" or "et al"
- Common types: [J] journal, [M] book, [C] conference, [D] thesis, [P] patent, [S] standard, [R] report, [J/OL] online journal

### Bilingual Requirements
- Title: Chinese + English
- Authors: Chinese + English (pinyin)
- Abstract: Chinese + English
- Keywords: Chinese + English
- Table/figure captions: Chinese + English
- Chinese references: Chinese + English

### Other Rules
- CLC number (中图分类号) and Document Code (文献标志码: A) should be present
- Equations: numbered (1), (2), (3)... in order of appearance
- Use MathType or Word equation editor (no image equations)
- SI units per GB 3100-1993 to GB 3102-1993
- Variable symbols in italic, units in upright
- Vectors/matrices in bold italic

## Scripts

### extract_rules.py
Reads a template .docx and outputs a JSON rules file with all detected formatting rules.

Usage: `python scripts/extract_rules.py <template.docx> [output.json]`

### format_docx.py
Applies formatting rules to a .docx file. Runs in phases (controlled by `--phase` flag).

Usage: `python scripts/format_docx.py <input.docx> <rules.json> --output <output.docx> [--phase A|B|C|D|E|all]`

### format_latex.py
Applies formatting rules to a .tex file.

Usage: `python scripts/format_latex.py <input.tex> <rules.json> --output <output.tex>`

### verify_format.py
Checks formatted document against rules and reports issues.

Usage: `python scripts/verify_format.py <formatted.docx> <rules.json>`

## Implementation Notes

### Detecting Chinese vs English text
```python
def is_chinese(char):
    return '一' <= char <= '鿿' or '　' <= char <= '〿'

def set_run_font(run, chinese_font, english_font, size_pt):
    """Set font for a run, handling CJK vs Latin."""
    run.font.size = Pt(size_pt)
    run.font.name = english_font
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), chinese_font)
    rFonts.set(qn('w:ascii'), english_font)
    rFonts.set(qn('w:hAnsi'), english_font)
```

### Three-line table
```python
def apply_three_line_table(table):
    """Convert a table to three-line academic style."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    # Remove all existing borders
    for tc in tbl.iter_tcs():
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'nil')
            tcBorders.append(border)
        tcPr.append(tcBorders)
```

## Common Mistakes to Avoid

1. **Don't just copy the template's visual font** — read the text rules (e.g., "正文叙述用5号宋体")
2. **Don't skip bilingual elements** — Chinese papers need English title, abstract, keywords, and translated references
3. **Don't blindly convert all periods** — English periods are kept in English context (after Latin letters); only periods after Chinese characters are converted to `。`. URLs, numbers, and abbreviations are auto-protected.
4. **Don't leave tables as Table Grid** — convert to three-line
5. **Don't number the introduction (引言)** — it should have no section number
6. **Don't use Chinese punctuation in English text** — and vice versa
7. **Don't forget to add "(in Chinese)"** after translated references
8. **Don't format in one shot without user checkpoints** — phase the work

## Red Flags — STOP and Ask User

- Template file cannot be read → ask user to verify path
- Target document structure is ambiguous → ask user to clarify which section is which
- More than 50% of rules cannot be automatically applied → present manual checklist instead
- python-docx not installed → install it first
