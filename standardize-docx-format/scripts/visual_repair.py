#!/usr/bin/env python3
"""Apply conservative visual repairs that do not rewrite document content."""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path
from typing import Any

from apply_profile import (
    PPR_ORDER,
    element_children,
    ensure_direct,
    parse_xml,
    set_toggle,
    validate_package_parts,
    write_package,
    xml_bytes,
)
from visual_qa import EMU_PER_INCH, visual_qa


SAFE_REPAIRS = {
    "orphan_heading",
    "orphan_heading_risk",
    "heading_body_split",
    "figure_caption_split",
    "table_caption_split",
    "image_overflow",
}


def iter_paragraphs(document) -> list:
    bodies = document.getElementsByTagName("w:body")
    if not bodies:
        return []
    out = []
    seen: set[int] = set()
    for block in element_children(bodies[0]):
        paragraphs = []
        if block.nodeName == "w:p":
            paragraphs.append(block)
        paragraphs.extend(list(block.getElementsByTagName("w:p")))
        for paragraph in paragraphs:
            identity = id(paragraph)
            if identity in seen:
                continue
            seen.add(identity)
            out.append(paragraph)
    return out


def set_keep_next(paragraph, enabled: bool = True) -> None:
    p_pr = ensure_direct(paragraph, "w:pPr")
    set_toggle(p_pr, "w:keepNext", enabled, PPR_ORDER)


def set_page_break_before(paragraph, enabled: bool = True) -> None:
    p_pr = ensure_direct(paragraph, "w:pPr")
    set_toggle(p_pr, "w:pageBreakBefore", enabled, PPR_ORDER)


def scale_drawings(paragraph, max_width_in: float) -> int:
    changed = 0
    max_cx = int(max_width_in * EMU_PER_INCH)
    extents = list(paragraph.getElementsByTagName("wp:extent"))
    extents.extend(paragraph.getElementsByTagName("a:ext"))
    seen: set[int] = set()
    for extent in extents:
        if id(extent) in seen:
            continue
        seen.add(id(extent))
        cx = extent.getAttribute("cx")
        cy = extent.getAttribute("cy")
        if not cx or not cx.isdigit():
            continue
        width = int(cx)
        if width <= max_cx:
            continue
        ratio = max_cx / width
        extent.setAttribute("cx", str(max_cx))
        if cy and cy.lstrip("-").isdigit():
            extent.setAttribute("cy", str(max(1, int(int(cy) * ratio))))
        changed += 1
    return changed


def apply_visual_repairs(
    input_path: Path,
    output_path: Path,
    issues: list[dict[str, Any]],
    force: bool = False,
) -> dict[str, Any]:
    if input_path.resolve() == output_path.resolve():
        raise ValueError("Refusing in-place write: input and output paths are identical")
    with zipfile.ZipFile(input_path, "r") as archive:
        source_infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in source_infos}
        corrupt = archive.testzip()
        if corrupt:
            raise ValueError(f"Input DOCX has corrupt member: {corrupt}")
    document = parse_xml(parts["word/document.xml"])
    paragraphs = iter_paragraphs(document)
    applied: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    keep_next_done: set[int] = set()

    for issue in issues:
        issue_type = issue.get("type")
        if issue_type not in SAFE_REPAIRS:
            skipped.append({"type": issue_type, "reason": "not a conservative visual repair"})
            continue
        index = issue.get("paragraphIndex")
        if issue_type == "image_overflow":
            width = issue.get("contentWidthIn")
            if not width:
                skipped.append({"type": issue_type, "reason": "missing content width"})
                continue
            changed = 0
            if isinstance(index, int) and 0 <= index < len(paragraphs):
                changed += scale_drawings(paragraphs[index], float(width))
            else:
                for paragraph in paragraphs:
                    changed += scale_drawings(paragraph, float(width))
            applied.append({"type": issue_type, "action": "scale_drawing", "changed": changed, "paragraphIndex": index})
            continue
        if not isinstance(index, int) or index < 0 or index >= len(paragraphs):
            skipped.append({"type": issue_type, "reason": "paragraph index out of range"})
            continue
        if issue.get("repair") == "set_page_break_before":
            set_page_break_before(paragraphs[index], True)
            applied.append({"type": issue_type, "action": "set_page_break_before", "paragraphIndex": index})
            continue
        if index in keep_next_done:
            continue
        set_keep_next(paragraphs[index], True)
        keep_next_done.add(index)
        applied.append({"type": issue_type, "action": "set_keep_next", "paragraphIndex": index})

    parts["word/document.xml"] = xml_bytes(document)
    validate_package_parts(parts)
    write_package(source_infos, parts, output_path, force)
    return {
        "ok": True,
        "input": str(input_path.resolve()),
        "output": str(output_path.resolve()),
        "applied": applied,
        "skipped": skipped,
        "appliedCount": len(applied),
        "skippedCount": len(skipped),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply conservative visual repairs to a new DOCX.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--visual", type=Path, help="Existing visual.json; otherwise run visual QA first")
    parser.add_argument("--profile", type=Path)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    try:
        if args.visual:
            visual = json.loads(args.visual.read_text(encoding="utf-8"))
        else:
            profile = json.loads(args.profile.read_text(encoding="utf-8")) if args.profile else None
            visual = visual_qa(args.input, profile, None)
        result = apply_visual_repairs(args.input, args.output, visual.get("issues") or [], args.force)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
