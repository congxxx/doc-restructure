# doc-restructure

Restructure Word documents (.docx) by migrating chapters, adjusting heading levels, renumbering figures/tables, and verifying integrity.

按照映射规则重组 Word 文档（.docx）章节，自动调整标题层级、图表编号，并验证完整性。

---

## Features / 功能

| Feature | Description |
|---|---|
| **Chapter migration** | Move sections across chapters according to a JSON mapping |
| **Heading level adjustment** | Automatically adjust H1-H4 levels to match the new structure |
| **Figure/table renumbering** | Renumber figures and tables sequentially per chapter, fix all in-text references |
| **Heading renumbering** | Renumber headings hierarchically (1.1, 1.1.1, 1.1.1.1, ...) |
| **Integrity verification** | Check table-caption correspondence, numbering continuity, and heading consistency |
| **Auto cleanup** | Remove temporary files after restructuring with `--clean` |

| 功能 | 说明 |
|---|---|
| **章节迁移** | 按 JSON 映射规则将原文档章节迁移到新结构 |
| **标题层级调整** | 自动调整 H1-H4 级别，保持层级正确 |
| **图表重编号** | 按新章节连续编号（图1-1, 图1-2, ...），同步修正正文引用 |
| **标题重编号** | 按层级连续编号（1.1, 1.1.1, 1.1.1.1, ...） |
| **完整性验证** | 检查表注与表格对应、图表编号连续性、标题编号连续性 |
| **自动清理** | 通过 `--clean` 参数在重构完成后自动删除临时文件 |

---

## Installation / 安装

**Requirements / 依赖：** Python >= 3.7, python-docx >= 0.8.11

```bash
# Install from source / 从源码安装
git clone https://github.com/congxxx/doc-restructure.git
cd doc-restructure
pip install -e .

# Or just install the dependency / 或仅安装依赖
pip install python-docx
```

---

## Quick Start / 快速开始

```bash
# 1. Analyze document structure / 分析文档结构
python -m doc_restructure analyze "input.docx"

# 2. Generate mapping template / 生成映射模板
python -m doc_restructure generate-mapping "input.docx" mapping.json
# Edit mapping.json to define the new structure / 编辑 mapping.json 定义新结构

# 3. Restructure / 执行重构
python -m doc_restructure restructure "input.docx" "output.docx" mapping.json --clean

# 4. Fix numbering / 修正编号
python -m doc_restructure renumber "output.docx"
python -m doc_restructure fix-headings "output.docx"

# 5. Verify / 验证
python -m doc_restructure verify "output.docx"
```

> **Note / 注意：** Run commands from the skill directory (where `doc_restructure/` is located), or install with `pip install -e .` first.
>
> 在 skill 目录（`doc_restructure/` 所在目录）下运行命令，或先通过 `pip install -e .` 安装。

---

## Commands / 命令

| Command / 命令 | Description / 说明 |
|---|---|
| `analyze <docx>` | Print document structure (headings, figures, tables) / 输出文档结构 |
| `generate-mapping <docx> [out.json]` | Auto-generate mapping template / 自动生成映射模板 |
| `restructure <in> <out> <map> [--clean]` | Restructure document / 重构文档；`--clean` 自动清理临时文件 |
| `renumber <docx>` | Renumber figures/tables per chapter / 按章节重编号图表 |
| `fix-headings <docx>` | Renumber headings hierarchically / 按层级重编号标题 |
| `verify <docx>` | Check integrity / 检查完整性 |
| `cleanup [dir]` | Remove temporary files / 删除临时文件 |

---

## Mapping Format / 映射格式

The mapping file defines the new document structure. Each chapter contains sections, and each section lists source headings to collect.

映射文件定义新文档结构。每个章节包含若干节，每节列出要收集的源标题。

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
        },
        {
          "title": "1.2 Key Challenges",
          "sources": [
            {"chapter": 4, "heading": "5.1 Identification", "target_level": 3}
          ]
        }
      ]
    }
  ]
}
```

### Fields / 字段说明

| Field / 字段 | Description / 说明 |
|---|---|
| `chapter` | Source chapter index (0-based) / 原文档章节索引（从 0 开始） |
| `heading` | Heading text to match (partial match supported) / 匹配的标题文本（支持部分匹配） |
| `target_level` | Heading level in new doc (2=H2, 3=H3, 4=H4) / 新文档中的标题级别 |

### Collection Scope / 收集范围

Starting from the specified heading, collect everything until the next heading at the **same or higher** level. Includes:
- The heading itself
- All sub-headings and body paragraphs
- Tables and images in between
- Trailing tables after figure/table captions (auto-extended)

从指定标题开始，到下一个**同级或更高级**标题为止。包含：
- 标题本身
- 所有子标题和正文段落
- 中间的表格和图片
- 表注/图注后的表格（自动扩展）

---

## Best Practices / 最佳实践

### Map to H2 level only / 只映射到 H2 级别

The tool automatically collects all child H3 content under each H2. Do NOT simultaneously map a parent H2 and its child H3 headings — the H3 entries will be silently skipped.

工具会自动收集 H2 下的所有 H3 内容。不要同时映射父标题和子标题——H3 条目会被静默跳过。

```json
// Correct / 正确：H3 children auto-included / H3 子标题自动包含
{"chapter": 0, "heading": "1.1 Project Understanding", "target_level": 3}

// Avoid / 避免：H3 is already in H2, will be skipped / H3 已在 H2 中，会被跳过
{"chapter": 0, "heading": "1.1 Project Understanding", "target_level": 3},
{"chapter": 0, "heading": "1.1.1 Background", "target_level": 4}  // skipped / 被跳过
```

### Use `--clean` for auto-cleanup / 使用 `--clean` 自动清理

Add `--clean` to the restructure command to automatically remove temporary files (mapping.json, headings.json, etc.) after successful restructuring.

在 restructure 命令后加 `--clean`，重构成功后自动删除临时文件。

```bash
python -m doc_restructure restructure input.docx output.docx mapping.json --clean
```

---

## Claude Code Integration / Claude Code 集成

This tool is designed as a Claude Code Skill. Place it in `.claude/skills/` and Claude will automatically use it when you ask to restructure a document.

本工具设计为 Claude Code Skill，放在 `.claude/skills/` 目录下，Claude 会在你要求重构文档时自动使用。

**Example conversation / 对话示例：**

> **User:** Help me restructure this document according to the evaluation requirements in 评审要求.md. Restructure into 4 chapters, do not change content, keep formatting consistent.
>
> **Claude:** [Analyzes document, reads requirements, generates mapping, executes restructuring, fixes numbering, verifies integrity]

---

## Technical Details / 技术细节

- **Style IDs are numbers / 样式 ID 是数字**: Word uses numeric style IDs internally (Heading 1 = "3"). The tool handles this automatically. / Word 内部使用数字样式 ID，工具自动处理。
- **Tables are separate elements / 表格是独立元素**: `doc.paragraphs` does not include tables. The tool uses `doc.element.body` to access all elements. / `doc.paragraphs` 不包含表格，工具通过 `doc.element.body` 访问所有元素。
- **Text spans multiple runs / 文本跨 Run 分布**: Word often splits text across runs. The tool handles cross-run replacement. / Word 常将文本拆分到多个 Run 中，工具支持跨 Run 替换。
- **Tables follow captions / 表格紧跟表注**: The tool auto-extends collection range to include trailing tables. / 工具自动扩展收集范围以包含表注后的表格。

---

## Known Limitations / 已知限制

| Limitation / 限制 | Details / 详情 |
|---|---|
| `.docx` only / 仅支持 .docx | Legacy `.doc` format is not supported / 不支持旧版 .doc 格式 |
| No header/footer refs / 不处理页眉页脚 | Figure/table references in headers/footers are not updated / 页眉页脚中的图表引用不会更新 |
| Formatting loss / 格式损失 | Cross-run text replacement may lose per-run formatting (colors, bold) / 跨 Run 替换可能丢失部分格式 |
| Image detection / 图片检测 | Some embedding styles (VML) may not be detected / 某些嵌入方式（VML）可能检测不到 |

---

## Project Structure / 项目结构

```
doc-restructure/
├── SKILL.md                  # Claude Code Skill definition / Skill 定义
├── README.md                 # This file / 本文件
├── LICENSE                   # MIT License
├── setup.py                  # Package setup / 包配置
├── requirements.txt          # Dependencies / 依赖
├── doc_restructure/          # Python package / Python 包
│   ├── __init__.py
│   ├── __main__.py           # CLI entry point / CLI 入口
│   ├── utils.py              # Shared utilities / 共享工具函数
│   ├── analyzer.py           # Document analysis / 文档分析
│   ├── restructurer.py       # Chapter migration (core) / 章节迁移（核心）
│   ├── renumberer.py         # Figure/table/heading renumbering / 图表标题编号
│   └── verifier.py           # Integrity verification / 完整性验证
└── examples/
    └── mapping.json          # Example mapping / 映射示例
```

---

## License / 许可证

[MIT](LICENSE)
