"""CLI entry point for doc-restructure."""

import sys
import json
from pathlib import Path

from .analyzer import analyze
from .restructurer import restructure
from .renumberer import renumber_figures, renumber_headings
from .verifier import verify

# Known temporary file patterns created during the workflow
TEMP_FILE_PATTERNS = [
    "mapping.json",
    "auto_mapping.json",
    "headings.json",
    "heading_list.txt",
    "heading_list_h3.txt",
    "verify_output.txt",
]


USAGE = """Usage: doc-restructure <command> [args]

Commands:
  analyze           <input.docx>                    Analyze document structure
  generate-mapping  <input.docx> [output.json]      Generate mapping template
  restructure       <input.docx> <output.docx> <mapping.json> [--clean]  Restructure document
  renumber          <docx>                          Renumber figures/tables
  fix-headings      <docx>                          Fix heading numbers
  verify            <docx>                          Verify document integrity
  cleanup           [directory]                     Remove temporary files

Options:
  --clean           Auto-remove temp files after successful restructure
"""


def generate_mapping(doc_path, output_path=None):
    """Generate a mapping JSON template from document structure.

    Creates a mapping where each H1 becomes a chapter and each H2 becomes a section.
    Users can then edit the mapping to reorganize chapters.

    Args:
        doc_path: Path to source .docx file
        output_path: Optional path to save JSON (prints to stdout if None)
    """
    from docx import Document
    doc = Document(doc_path)
    paragraphs = doc.paragraphs

    chapters = []
    current_chapter = None
    h1_idx = 0

    for p in paragraphs:
        if not p.style:
            continue
        if p.style.name == "Heading 1":
            current_chapter = {
                "title": p.text.strip(),
                "sections": []
            }
            chapters.append(current_chapter)
            h1_idx += 1
        elif p.style.name == "Heading 2" and current_chapter is not None:
            section = {
                "title": p.text.strip(),
                "sources": [
                    {
                        "chapter": h1_idx - 1,
                        "heading": p.text.strip(),
                        "target_level": 3
                    }
                ]
            }
            current_chapter["sections"].append(section)

    mapping = {"chapters": chapters}

    output = json.dumps(mapping, ensure_ascii=False, indent=2)
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"Generated mapping: {output_path}")
        print(f"Chapters: {len(chapters)}")
        total_sections = sum(len(c['sections']) for c in chapters)
        print(f"Sections: {total_sections}")
        print("\nEdit the mapping file to reorganize chapters, then run:")
        print(f"  python -m doc_restructure restructure <input.docx> <output.docx> {output_path}")
    else:
        print(output)


def cleanup(directory="."):
    """Remove temporary files created during the workflow.

    Args:
        directory: Directory to clean (default: current directory)

    Returns:
        Number of files removed
    """
    directory = Path(directory)
    removed = 0
    for pattern in TEMP_FILE_PATTERNS:
        f = directory / pattern
        if f.exists():
            f.unlink()
            print(f"  Removed: {f.name}")
            removed += 1
    if removed == 0:
        print("  No temporary files found.")
    else:
        print(f"\nCleaned up {removed} temporary file(s).")
    return removed


def main():
    if len(sys.argv) < 2:
        print(USAGE)
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == "analyze":
        if len(sys.argv) < 3:
            print("Usage: doc-restructure analyze <docx>")
            sys.exit(1)
        analyze(sys.argv[2])

    elif cmd == "generate-mapping":
        if len(sys.argv) < 3:
            print("Usage: doc-restructure generate-mapping <docx> [output.json]")
            sys.exit(1)
        output = sys.argv[3] if len(sys.argv) > 3 else None
        generate_mapping(sys.argv[2], output)

    elif cmd == "restructure":
        if len(sys.argv) < 5:
            print("Usage: doc-restructure restructure <input> <output> <mapping.json> [--clean]")
            sys.exit(1)
        mapping_path = sys.argv[4]
        with open(mapping_path, "r", encoding="utf-8") as f:
            mapping = json.load(f)
        restructure(sys.argv[2], sys.argv[3], mapping)

        # Auto-cleanup if --clean flag is present
        if "--clean" in sys.argv:
            mapping_dir = Path(mapping_path).parent
            print(f"\n=== Cleanup ===")
            cleanup(mapping_dir)

    elif cmd == "renumber":
        if len(sys.argv) < 3:
            print("Usage: doc-restructure renumber <docx>")
            sys.exit(1)
        renumber_figures(sys.argv[2])

    elif cmd == "fix-headings":
        if len(sys.argv) < 3:
            print("Usage: doc-restructure fix-headings <docx>")
            sys.exit(1)
        renumber_headings(sys.argv[2])

    elif cmd == "verify":
        if len(sys.argv) < 3:
            print("Usage: doc-restructure verify <docx>")
            sys.exit(1)
        verify(sys.argv[2])

    elif cmd == "cleanup":
        directory = sys.argv[2] if len(sys.argv) > 2 else "."
        cleanup(directory)

    else:
        print(f"Unknown command: {cmd}")
        print(USAGE)
        sys.exit(1)


if __name__ == "__main__":
    main()
