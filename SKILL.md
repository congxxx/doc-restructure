---
name: doc-restructure
description: >
  Restructure Word document (.docx) chapters according to mapping rules.
  Handles chapter migration, heading level adjustment, figure/table renumbering,
  cross-reference fixing, and integrity verification.
  Use this skill when the user mentions: restructuring a Word document,
  reorganizing document chapters, merging/splitting chapters, renumbering
  figures or tables, fixing heading numbers, or adapting a document to
  match evaluation requirements. Also use when the user has a multi-chapter
  .docx that needs to be reorganized into a different structure.
---

# doc-restructure

Restructure Word documents by migrating sections, adjusting heading levels, renumbering figures/tables, and verifying integrity.

## Quick Start

The Python package is bundled with this skill at `doc_restructure/`.

```bash
# Install dependency (one-time)
pip install python-docx

# 1. Analyze the source document
python -m doc_restructure analyze "input.docx"

# 2. Auto-generate mapping template, then edit it
python -m doc_restructure generate-mapping "input.docx" mapping.json

# 3. Restructure (--clean auto-removes temp files after success)
python -m doc_restructure restructure "input.docx" "output.docx" mapping.json --clean

# 4. Fix numbering
python -m doc_restructure renumber "output.docx"
python -m doc_restructure fix-headings "output.docx"

# 5. Verify
python -m doc_restructure verify "output.docx"
```

**Important**: Run all commands from the skill directory (where `doc_restructure/` is located), or install the package first with `pip install -e .`.

## Workflow

When the user asks to restructure a document:

1. **Analyze** the source document first — understand its heading structure, figure/table distribution, and chapter boundaries.
2. **Read the requirements** — if the user provides evaluation criteria or target structure, read it to understand the desired chapters.
3. **Generate mapping.json** — create the mapping that defines the new structure. Present it to the user for confirmation before executing.
4. **Execute** the restructuring.
5. **Fix numbering** — renumber figures/tables and headings.
6. **Verify** — run verify to confirm nothing is missing.

## Mapping Format

```json
{
  "chapters": [
    {
      "title": "一、Chapter Title",
      "sections": [
        {
          "title": "1.1 Section Title",
          "sources": [
            {"chapter": 0, "heading": "1.1.1 Original Heading", "target_level": 3},
            {"chapter": 0, "heading": "1.2 Another Heading", "target_level": 3}
          ]
        }
      ]
    }
  ]
}
```

- `chapter`: Source document chapter index (0-based, first chapter = 0)
- `heading`: Text to match in source headings (partial match supported)
- `target_level`: Heading level in new document (2=H2, 3=H3, 4=H4)

## Key Technical Details

These are important for understanding how the tool works:

- **Style IDs are numbers**: Word uses numeric style IDs internally (Heading 1 = "3", not "Heading 1"). The tool handles this automatically.
- **Tables are separate elements**: `doc.paragraphs` does not include tables. The tool uses `doc.element.body` to access all elements including tables.
- **Text spans multiple runs**: Word often splits text like "图1-2" across runs ("图1-" in one, "2" in another). The tool handles cross-run replacement.
- **Tables follow captions**: Table captions (表注) are paragraphs, but the actual table is a separate body element right after. The tool automatically extends the collection range to include trailing tables.

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

## Mapping Best Practices

**Map to H2 level only.** The tool automatically collects all child H3 content under each H2. Do NOT simultaneously map a parent H2 and its child H3 headings — the H3 entries will be silently skipped (they're already included in the H2).

```json
// Correct: map H2, H3 children are auto-included
{"chapter": 0, "heading": "1.1 项目理解", "target_level": 3}

// Avoid: mapping both H2 and its H3 children causes redundant entries
{"chapter": 0, "heading": "1.1 项目理解", "target_level": 3},
{"chapter": 0, "heading": "1.1.1 项目背景", "target_level": 4}  // silently skipped
```

## Known Limitations

- Only supports `.docx` format (not legacy `.doc`)
- Does not process figure/table references in headers/footers
- Cross-run text replacement may lose per-run formatting (colors, bold)
- Image detection is XML-based; some embedding styles may not be detected
