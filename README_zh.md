# 论文排版 Skill (Paper Formatting)

Claude Code 学术论文排版技能。读取期刊/学校投稿模板中的排版规范，自动将论文草稿格式化为符合要求的文档。支持 Word (.docx) 和 LaTeX (.tex)。

## 功能

- 读取模板中的**排版规则**（而非简单复制视觉格式）
- 自动设置字体、字号、页边距、行距、缩进
- 标题层级（一/二/三级标题不同字体字号）
- 表格转三线表
- 中英文标点修正（句号改"．"、逗号改"，"等）
- 参考文献顺序编码修正
- 双语元素缺失提醒（英文摘要/关键词/参考文献对照）
- 分阶段格式化，每阶段用户确认

## 安装

将整个文件夹复制到 `~/.claude/skills/paper-formatting/`：

```bash
git clone https://github.com/HH-ANTENNA/paper-formatting.git ~/.claude/skills/paper-formatting/
```

## 使用

在 Claude Code 中说：**"帮我排版论文"** 或 **"format my paper"**

Skill 会自动加载并引导：

1. **确认模板** — 「使用默认模板？（深大学报理工版）」
2. **提取规则** — 从模板文字说明中提取排版规范，展示给你确认
3. **分析文档** — 检测标题/作者/摘要/章节/图表/参考文献
4. **分阶段格式化**（每阶段用户确认）：

| 阶段 | 内容 |
|------|------|
| A | 页面设置 + 正文字体 + 段落 + 标点修正 |
| B | 前置部分（标题/作者/摘要/关键词） |
| C | 标题层级（一级/二级/三级） |
| D | 表格三线表 + 图题注 |
| E | 参考文献顺序编码 |

5. **验证报告** — 列出 ✅ 已修复项 + ⚠️ 需手动处理项

## 默认模板规范

默认使用**《深圳大学学报理工版》**投稿模板：

| 元素 | 字体 | 字号 |
|------|------|------|
| 正文 | 宋体 / Times New Roman | 五号 (10.5pt) |
| 一级标题 | 仿宋 | 三号 (16pt) |
| 二级标题 | 黑体 | 五号 (10.5pt) |
| 三级标题 | 楷体 | 五号 (10.5pt) |
| 作者 | 楷体 | 四号 (14pt) |
| 摘要/关键词 | 楷体 | 五号 (10.5pt) |
| 表题/图题 | 宋体 | 小五 (9pt) |
| 参考文献正文 | 宋体 | 小五 (9pt) |

## 脚本说明

| 脚本 | 功能 |
|------|------|
| `scripts/extract_rules.py` | 从模板 .docx 提取排版规则 → JSON |
| `scripts/format_docx.py` | Word 文档分阶段格式化 |
| `scripts/format_latex.py` | LaTeX 文档格式化 |
| `scripts/verify_format.py` | 格式化后验证检查 |

### 命令行用法

```bash
# 从模板提取规则
python scripts/extract_rules.py 模板.docx rules.json

# 使用内置默认规则
python scripts/extract_rules.py --default rules.json

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
1. Skill 问「使用默认模板？」时答 **否**
2. 提供你的 .docx 模板路径
3. Skill 会读取模板中的文字规则和示例格式

## License

MIT
