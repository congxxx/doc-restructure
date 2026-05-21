<div align="right">

[English](#english) | [中文](#中文)

</div>

<a id="english"></a>

# doc-restructure

Restructure Word documents (.docx) by migrating chapters, adjusting heading levels, renumbering figures/tables, and verifying integrity. Designed as a Claude Code Skill — just describe what you want and Claude handles the rest.

<div align="right"><a href="#中文">切换到中文</a></div>

---

## Features

- **Chapter migration** — Move sections across chapters according to a JSON mapping
- **Heading level adjustment** — Automatically adjust H1–H4 levels to match the new structure
- **Figure/table renumbering** — Renumber sequentially per chapter, fix all in-text references
- **Heading renumbering** — Renumber hierarchically (1.1, 1.1.1, 1.1.1.1, …)
- **Integrity verification** — Check table-caption correspondence, numbering continuity
- **Auto cleanup** — Remove temporary files after restructuring with `--clean`

## Installation

**Requirements:** Python >= 3.7, python-docx >= 0.8.11

```bash
git clone https://github.com/congxxx/doc-restructure.git
cd doc-restructure
pip install -e .
```

## Quick Start

```bash
# 1. Analyze document structure
python -m doc_restructure analyze "input.docx"

# 2. Generate mapping template, then edit it
python -m doc_restructure generate-mapping "input.docx" mapping.json

# 3. Restructure (--clean auto-removes temp files)
python -m doc_restructure restructure "input.docx" "output.docx" mapping.json --clean

# 4. Fix numbering
python -m doc_restructure renumber "output.docx"
python -m doc_restructure fix-headings "output.docx"

# 5. Verify
python -m doc_restructure verify "output.docx"
```

> Run commands from the skill directory (where `doc_restructure/` is located), or install with `pip install -e .` first.

## Commands

| Command | Description |
|---|---|
| `analyze <docx>` | Print document structure (headings, figures, tables) |
| `generate-mapping <docx> [out.json]` | Auto-generate mapping template from document structure |
| `restructure <in> <out> <map> [--clean]` | Restructure document; `--clean` auto-removes temp files |
| `renumber <docx>` | Renumber figures/tables sequentially per chapter |
| `fix-headings <docx>` | Renumber headings hierarchically (1.1, 1.1.1, etc.) |
| `verify <docx>` | Check table-caption correspondence, numbering continuity |
| `cleanup [dir]` | Remove temporary files (mapping.json, headings.json, etc.) |

## Mapping Format

The mapping file defines the new document structure. Each chapter contains sections, and each section lists source headings to collect.

```json
{
  "chapters": [
    {
      "title": "一、Project Technical Plan",
      "sections": [
        {
          "title": "1.1 Project Overview",
          "sources": [
            {"chapter": 0, "heading": "1.1 Background", "target_level": 3},
            {"chapter": 0, "heading": "1.2 Scope", "target_level": 3}
          ]
        }
      ]
    }
  ]
}
```

| Field | Description |
|---|---|
| `chapter` | Source chapter index (0-based) |
| `heading` | Heading text to match (partial match supported) |
| `target_level` | Heading level in new doc (2=H2, 3=H3, 4=H4) |

**Collection scope:** From the specified heading, collect everything until the next heading at the same or higher level — including sub-headings, body paragraphs, tables, images, and trailing tables after captions.

### Best Practice

**Map to H2 level only.** The tool automatically collects all child H3 content under each H2. Do NOT simultaneously map a parent H2 and its child H3 headings — the H3 entries will be silently skipped.

```json
// Correct: H3 children auto-included
{"chapter": 0, "heading": "1.1 Project Understanding", "target_level": 3}

// Avoid: H3 already in H2, will be skipped
{"chapter": 0, "heading": "1.1 Project Understanding", "target_level": 3},
{"chapter": 0, "heading": "1.1.1 Background", "target_level": 4}  // skipped
```

## Claude Code Integration

This tool works as a Claude Code Skill. Place it in `.claude/skills/` and Claude will automatically use it when you ask to restructure a document.

> **User:** Help me restructure this document according to the evaluation requirements. Restructure into 4 chapters, do not change content, keep formatting consistent.
>
> **Claude:** [Analyzes document → reads requirements → generates mapping → executes restructuring → fixes numbering → verifies integrity]

## Technical Details

- **Style IDs are numbers** — Word uses numeric style IDs internally (Heading 1 = "3"). The tool handles this automatically.
- **Tables are separate elements** — `doc.paragraphs` does not include tables. The tool uses `doc.element.body` to access all elements.
- **Text spans multiple runs** — Word often splits text across runs. The tool handles cross-run replacement.
- **Tables follow captions** — The tool auto-extends collection range to include trailing tables.

## Known Limitations

- Only supports `.docx` format (not legacy `.doc`)
- Does not update figure/table references in headers/footers
- Cross-run text replacement may lose per-run formatting (colors, bold)
- Image detection is XML-based; some embedding styles (VML) may not be detected

## Project Structure

```
doc-restructure/
├── SKILL.md                  # Claude Code Skill definition
├── README.md                 # This file
├── LICENSE                   # MIT
├── setup.py
├── requirements.txt
├── doc_restructure/          # Python package
│   ├── __init__.py
│   ├── __main__.py           # CLI entry point
│   ├── utils.py              # Shared utilities
│   ├── analyzer.py           # Document analysis
│   ├── restructurer.py       # Chapter migration (core)
│   ├── renumberer.py         # Figure/table/heading renumbering
│   └── verifier.py           # Integrity verification
└── examples/
    └── mapping.json          # Example mapping
```

## License

[MIT](LICENSE)

---

<a id="中文"></a>

# doc-restructure

按映射规则重组 Word 文档（.docx）章节，自动调整标题层级、图表编号，并验证完整性。设计为 Claude Code Skill——只需描述需求，Claude 自动完成。

<div align="right"><a href="#english">Switch to English</a></div>

---

## 功能

- **章节迁移** — 按 JSON 映射规则将原文档章节迁移到新结构
- **标题层级调整** — 自动调整 H1–H4 级别，保持层级正确
- **图表重编号** — 按新章节连续编号（图1-1, 图1-2, …），同步修正正文引用
- **标题重编号** — 按层级连续编号（1.1, 1.1.1, 1.1.1.1, …）
- **完整性验证** — 检查表注与表格对应、图表编号连续性
- **自动清理** — 通过 `--clean` 参数在重构完成后自动删除临时文件

## 安装

**依赖：** Python >= 3.7, python-docx >= 0.8.11

```bash
git clone https://github.com/congxxx/doc-restructure.git
cd doc-restructure
pip install -e .
```

## 快速开始

```bash
# 1. 分析文档结构
python -m doc_restructure analyze "input.docx"

# 2. 生成映射模板，然后编辑
python -m doc_restructure generate-mapping "input.docx" mapping.json

# 3. 执行重构（--clean 自动清理临时文件）
python -m doc_restructure restructure "input.docx" "output.docx" mapping.json --clean

# 4. 修正编号
python -m doc_restructure renumber "output.docx"
python -m doc_restructure fix-headings "output.docx"

# 5. 验证
python -m doc_restructure verify "output.docx"
```

> 在 skill 目录（`doc_restructure/` 所在目录）下运行命令，或先通过 `pip install -e .` 安装。

## 命令

| 命令 | 说明 |
|---|---|
| `analyze <docx>` | 输出文档结构（标题、图表） |
| `generate-mapping <docx> [out.json]` | 从文档结构自动生成映射模板 |
| `restructure <in> <out> <map> [--clean]` | 重构文档；`--clean` 自动清理临时文件 |
| `renumber <docx>` | 按章节重编号图表 |
| `fix-headings <docx>` | 按层级重编号标题 |
| `verify <docx>` | 检查完整性 |
| `cleanup [dir]` | 删除临时文件 |

## 映射格式

映射文件定义新文档结构。每个章节包含若干节，每节列出要收集的源标题。

```json
{
  "chapters": [
    {
      "title": "一、项目技术服务方案",
      "sections": [
        {
          "title": "1.1 项目概况分析",
          "sources": [
            {"chapter": 0, "heading": "1.1 项目理解", "target_level": 3},
            {"chapter": 0, "heading": "1.2 项目区概况", "target_level": 3}
          ]
        }
      ]
    }
  ]
}
```

| 字段 | 说明 |
|---|---|
| `chapter` | 原文档章节索引（从 0 开始） |
| `heading` | 匹配的标题文本（支持部分匹配） |
| `target_level` | 新文档中的标题级别（2=H2, 3=H3, 4=H4） |

**收集范围：** 从指定标题开始，到下一个同级或更高级标题为止——包含子标题、正文段落、表格、图片，以及表注/图注后的表格（自动扩展）。

### 最佳实践

**只映射到 H2 级别。** 工具会自动收集 H2 下的所有 H3 内容。不要同时映射父标题和子标题——H3 条目会被静默跳过。

```json
// 正确：H3 子标题自动包含
{"chapter": 0, "heading": "1.1 项目理解", "target_level": 3}

// 避免：H3 已在 H2 中，会被跳过
{"chapter": 0, "heading": "1.1 项目理解", "target_level": 3},
{"chapter": 0, "heading": "1.1.1 项目背景", "target_level": 4}  // 被跳过
```

## Claude Code 集成

本工具设计为 Claude Code Skill，放在 `.claude/skills/` 目录下，Claude 会在你要求重构文档时自动使用。

> **用户：** 帮我按照评审要求重构这个文档，重构为四个章节，不要改变内容，格式保持一致。
>
> **Claude：** [分析文档 → 读取要求 → 生成映射 → 执行重构 → 修正编号 → 验证完整性]

## 技术细节

- **样式 ID 是数字** — Word 内部使用数字样式 ID（Heading 1 = "3"），工具自动处理
- **表格是独立元素** — `doc.paragraphs` 不包含表格，工具通过 `doc.element.body` 访问所有元素
- **文本跨 Run 分布** — Word 常将文本拆分到多个 Run 中，工具支持跨 Run 替换
- **表格紧跟表注** — 工具自动扩展收集范围以包含表注后的表格

## 已知限制

- 仅支持 `.docx` 格式（不支持旧版 `.doc`）
- 不更新页眉页脚中的图表引用
- 跨 Run 文本替换可能丢失部分格式（颜色、加粗）
- 图片检测基于 XML，某些嵌入方式（VML）可能检测不到

## 项目结构

```
doc-restructure/
├── SKILL.md                  # Claude Code Skill 定义
├── README.md                 # 本文件
├── LICENSE                   # MIT
├── setup.py
├── requirements.txt
├── doc_restructure/          # Python 包
│   ├── __init__.py
│   ├── __main__.py           # CLI 入口
│   ├── utils.py              # 共享工具函数
│   ├── analyzer.py           # 文档分析
│   ├── restructurer.py       # 章节迁移（核心）
│   ├── renumberer.py         # 图表/标题编号
│   └── verifier.py           # 完整性验证
└── examples/
    └── mapping.json          # 映射示例
```

## 许可证

[MIT](LICENSE)
