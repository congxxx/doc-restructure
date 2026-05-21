"""Document integrity verifier."""

import re
from docx import Document
from docx.oxml.ns import qn
from .utils import detect_style_mapping, get_heading_level, get_elem_tag, get_body_elements


def verify(doc_path):
    """Verify document integrity and print report.

    Checks:
    1. Table captions followed by tables
    2. Figure captions with images
    3. Figure/table numbering continuity
    4. Heading numbering continuity

    Args:
        doc_path: Path to .docx file
    """
    doc = Document(doc_path)
    paragraphs = doc.paragraphs
    body_elements = get_body_elements(doc)

    chapter_starts = [
        i for i, p in enumerate(paragraphs)
        if p.style and p.style.name == "Heading 1"
    ]

    all_issues = []

    # Single-pass body element check for table/figure captions
    print("=== Captions & Images ===")
    table_ok, table_missing, missing_captions, fig_ok, fig_missing = \
        _check_captions_single_pass(body_elements)
    print(f"  Tables: {table_ok} OK, {table_missing} missing")
    print(f"  Figures: {fig_ok} OK, {fig_missing} possibly missing")
    all_issues.extend(missing_captions)

    # 3. Check figure/table numbering continuity
    print("\n=== Figure/Table Numbering ===")
    fig_re = re.compile(r"图\s*(\d+)\s*-\s*(\d+)")
    tab_re = re.compile(r"表\s*(\d+)\s*-\s*(\d+)")

    for ci in range(len(chapter_starts)):
        s = chapter_starts[ci]
        e = chapter_starts[ci + 1] if ci + 1 < len(chapter_starts) else len(paragraphs)
        ch_num = ci + 1

        figs = []
        tabs = []
        for i in range(s, e):
            p = paragraphs[i]
            if p.style and "图注" in p.style.name:
                m = fig_re.search(p.text)
                if m:
                    figs.append((int(m.group(1)), int(m.group(2))))
            elif p.style and "表注" in p.style.name:
                m = tab_re.search(p.text)
                if m:
                    tabs.append((int(m.group(1)), int(m.group(2))))

        issues = []
        expected = 1
        for ch, num in figs:
            if ch != ch_num or num != expected:
                issues.append(f"图{ch}-{num} (should be 图{ch_num}-{expected})")
            expected = num + 1

        expected = 1
        for ch, num in tabs:
            if ch != ch_num or num != expected:
                issues.append(f"表{ch}-{num} (should be 表{ch_num}-{expected})")
            expected = num + 1

        if issues:
            all_issues.extend(issues)
            print(f"  Chapter {ch_num}: {len(issues)} issues")
            for iss in issues[:3]:
                print(f"    {iss}")
        else:
            print(f"  Chapter {ch_num}: {len(figs)} figs, {len(tabs)} tables - OK")

    # 4. Check heading numbering continuity
    print("\n=== Heading Numbering ===")
    for ci in range(len(chapter_starts)):
        s = chapter_starts[ci]
        e = chapter_starts[ci + 1] if ci + 1 < len(chapter_starts) else len(paragraphs)
        ch_num = ci + 1

        h2_expected = 1
        h3_expected = 1
        h_issues = []

        for i in range(s, e):
            p = paragraphs[i]
            level = get_heading_level(p)
            if level == 2:
                m = re.match(r"^(\d+)\.(\d+)", p.text.strip())
                if m:
                    if int(m.group(1)) != ch_num or int(m.group(2)) != h2_expected:
                        h_issues.append(
                            f"H2 '{p.text.strip()[:20]}' (should be {ch_num}.{h2_expected})"
                        )
                h2_expected += 1
                h3_expected = 1
            elif level == 3:
                m = re.match(r"^(\d+)\.(\d+)\.(\d+)", p.text.strip())
                if m:
                    expected_h2 = h2_expected - 1
                    if (
                        int(m.group(1)) != ch_num
                        or int(m.group(2)) != expected_h2
                        or int(m.group(3)) != h3_expected
                    ):
                        h_issues.append(
                            f"H3 '{p.text.strip()[:20]}' "
                            f"(should be {ch_num}.{expected_h2}.{h3_expected})"
                        )
                h3_expected += 1

        if h_issues:
            all_issues.extend(h_issues)
            print(f"  Chapter {ch_num}: {len(h_issues)} issues")
        else:
            print(f"  Chapter {ch_num}: OK")

    # Summary
    print("\n=== Summary ===")
    if all_issues:
        print(f"Found {len(all_issues)} issues:")
        for iss in all_issues[:10]:
            print(f"  - {iss}")
        if len(all_issues) > 10:
            print(f"  ... and {len(all_issues) - 10} more")
    else:
        print("No issues found!")

    return len(all_issues) == 0


def _check_captions_single_pass(body_elements):
    """Check table and figure captions in a single pass over body elements.

    Returns:
        (table_ok, table_missing, missing_captions, fig_ok, fig_missing)
    """
    table_ok = 0
    table_missing = 0
    missing_captions = []
    fig_ok = 0
    fig_missing = 0

    drawing_tag = qn("w:drawing")
    val_attr = qn("w:val")

    n = len(body_elements)
    for i in range(n):
        elem = body_elements[i]
        tag = elem.tag
        # Fast path: skip non-paragraph elements
        if not tag.endswith("}p"):
            continue

        pPr = elem.find(qn("w:pPr"))
        if pPr is None:
            continue
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None:
            continue

        style_val = pStyle.get(val_attr)
        if style_val is None:
            continue

        if style_val == "178":
            # Table caption
            if i + 1 < n and body_elements[i + 1].tag.endswith("}tbl"):
                table_ok += 1
            else:
                table_missing += 1
                caption = _get_elem_text(elem)[:50]
                missing_captions.append(f"Table missing: {caption}")

        elif style_val == "176":
            # Figure caption
            has_image = _has_drawing(elem, drawing_tag)
            if not has_image and i + 1 < n:
                has_image = _has_drawing(body_elements[i + 1], drawing_tag)
            if has_image:
                fig_ok += 1
            else:
                fig_missing += 1

    return table_ok, table_missing, missing_captions, fig_ok, fig_missing


def _has_drawing(elem, drawing_tag):
    """Check if element contains a drawing (image)."""
    for r in elem.iter(qn("w:r")):
        if r.find(".//" + drawing_tag) is not None:
            return True
    return False


def _get_elem_text(elem):
    """Extract text from a paragraph element."""
    texts = []
    for t in elem.iter(qn("w:t")):
        if t.text:
            texts.append(t.text)
    return "".join(texts)
