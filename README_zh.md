# 论文排版 Skill (Paper Formatting)

[🇺🇸 English](README.md)

Claude Code 学术论文排版技能。读取期刊/学校投稿模板中的排版规范，自动将论文草稿格式化为符合要求的文档。支持 Word (.docx) 和 LaTeX (.tex)。

## ⚠️ 免责声明

**本技能仅供学习交流参考，不构成任何形式的格式担保。** 虽然可以自动完成常见排版操作，但无法覆盖所有边界情况。不同学校、期刊、审稿人对格式要求可能有所不同。**请务必逐项人工核对排版结果**——检查字体字号、标题层级、表格样式、参考文献格式、页面布局等是否符合官方规范。因使用本技能导致的格式问题，作者不承担任何责任。

## 功能

- 读取模板排版规则，自动应用到草稿
- 自动设置字体、字号、页边距、行距、缩进
- 标题层级（黑体加粗，三号→小三→四号逐级递减），应用 Word 内置 Heading 样式（导航窗格可折叠/展开）
- 摘要/关键词使用【】黑色方头括号（U+3010/U+3011）内联格式
- 上下文感知中英文标点修正（中文标点 → 全角，英文标点保持不变）
- 表格转三线表
- 参考文献顺序编码 + 悬挂缩进
- 分阶段格式化，每阶段用户确认

## 安装

将整个文件夹复制到 `~/.claude/skills/paper-formatting/`：

```bash
git clone https://github.com/HH-ANTENNA/paper-formatting.git ~/.claude/skills/paper-formatting/
```

## 使用

在 Claude Code 中说：**"帮我排版论文"** 或 **"format my paper"**

Skill 会自动加载并引导：

1. **生成规则** — 使用默认规则或提取自定义模板规则
2. **分析文档** — 检测标题/摘要/关键词/章节/图表/参考文献
3. **分阶段格式化**（每阶段用户确认）：

| 阶段 | 内容 |
|------|------|
| A | 页面设置 + 正文字体 + 段落间距 + 标点修正 |
| B | 前置部分（标题/【摘要】/【关键词】） |
| C | 标题层级（一级/二级/三级 + Heading 样式） |
| D | 表格三线表 + 图题注 |
| E | 参考文献顺序编码 + 悬挂缩进 |

4. **验证报告** — 列出 ✅ 已修复项 + ⚠️ 警告 + ❌ 需手动处理项

## 默认排版规范

| 元素 | 字体 | 字号 | 段落 |
|------|------|------|------|
| 标题 | 宋体 | 小二 (18pt) Bold | 居中 |
| 正文 | 宋体 / Times New Roman | 五号 (10.5pt) | 单倍行距, 段前段后0.5行 |
| 一级标题 "1  XXX" | 黑体 Bold | 三号 (16pt) | 单倍行距, 段前段后0.5行 |
| 二级标题 "1.1  XXX" | 黑体 Bold | 小三 (15pt) | 单倍行距, 段前段后0.5行 |
| 三级标题 "1.1.1  XXX" | 黑体 Bold | 四号 (14pt) | 单倍行距, 段前段后0.5行 |
| 【摘要】标签 | 楷体 | 小四 (12pt) | — |
| 摘要正文 | 楷体 | 五号 (10.5pt) | — |
| 【关键词】标签 | 楷体 Bold | 小四 (12pt) | — |
| 关键词条目 | 楷体 | 五号 (10.5pt) | — |
| 图题/表题 | 宋体 | 小五 (9pt) | 居中 |
| 参考文献标题 | 楷体 | 五号 (10.5pt) | — |
| 参考文献正文 | 宋体 | 小五 (9pt) | 悬挂缩进 |

- **页面：** A4, 页边距 上/下/左/右 = 2.54/2.54/3.17/3.17 cm
- **所有章节均编号**（含引言、参考文献）
- **标题** 全部使用 Word 内置 Heading 1/2/3 样式（导航窗格可折叠/展开）
- **摘要/关键词** 使用【】黑色方头括号，与正文同一段落（内联格式）
- **关键词** 用全角分号 `；` 分隔

## 脚本说明

| 脚本 | 功能 |
|------|------|
| `scripts/extract_rules.py` | 生成默认排版规则 JSON，或从模板 .docx 提取规则 |
| `scripts/format_docx.py` | Word 文档分阶段格式化 |
| `scripts/format_latex.py` | LaTeX 文档格式化 |
| `scripts/verify_format.py` | 格式化后验证检查 |

### 命令行用法

```bash
# 生成默认规则
python scripts/extract_rules.py --default rules.json

# 从自定义模板提取规则
python scripts/extract_rules.py 模板.docx rules.json

# 格式化文档（全部阶段）
python scripts/format_docx.py 论文.docx rules.json -o 格式化后.docx --phase all

# 仅执行指定阶段
python scripts/format_docx.py 论文.docx rules.json -o 格式化后.docx --phase A,B,C

# 验证格式化结果
python scripts/verify_format.py 格式化后.docx rules.json
```

## 环境要求

```bash
pip install python-docx
```

## 自定义模板

使用自己的模板：
1. 提供你的 .docx 模板路径
2. Skill 会读取模板中的文字规则和示例格式，覆盖默认规则

## License

MIT
