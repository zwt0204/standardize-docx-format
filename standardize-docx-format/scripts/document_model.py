#!/usr/bin/env python3
"""Build a semantic Document AST from a DOCX without mutating the file."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any
from xml.dom import Node

from audit_docx import (
    attr,
    audit,
    direct_child,
    element_children,
    parse_xml,
    paragraph_props,
    run_props,
)


HEADING_STYLE_IDS = {
    "Heading1": 1,
    "Heading2": 2,
    "Heading3": 3,
    "Heading4": 4,
    "Heading5": 5,
    "1": 1,
    "2": 2,
    "3": 3,
    "4": 4,
    "5": 5,
}

REGION_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    ("cover", re.compile(r"^(封面|题目|毕业论文|毕业设计|学位论文)$")),
    ("declaration", re.compile(r"^(独创性声明|诚信声明|版权使用授权书)")),
    ("abstract_zh", re.compile(r"^摘\s*要$")),
    ("abstract_en", re.compile(r"^abstract$", re.I)),
    ("toc", re.compile(r"^(目\s*录|contents)$", re.I)),
    ("list_of_figures", re.compile(r"^(图目录|插图目录|list of figures)$", re.I)),
    ("list_of_tables", re.compile(r"^(表目录|表格目录|list of tables)$", re.I)),
    ("references", re.compile(r"^(参考文献|reference[s]?)$", re.I)),
    ("acknowledgement", re.compile(r"^(致\s*谢|acknowledgements?)$", re.I)),
    ("appendix", re.compile(r"^(附录|appendix)\b", re.I)),
]

CHAPTER_PATTERNS = [
    re.compile(r"^第[一二三四五六七八九十百零0-9]+章"),
    re.compile(r"^第[0-9]+章"),
    re.compile(r"^[一二三四五六七八九十]+、"),
    re.compile(r"^\d+\s+\S+"),
]
SECTION_PATTERNS = [
    re.compile(r"^（[一二三四五六七八九十]+）"),
    re.compile(r"^\d+\.\d+\s+\S+"),
]
SUBSECTION_PATTERNS = [
    re.compile(r"^\d+\.\d+\.\d+\s+\S+"),
    re.compile(r"^[（(]\d+[）)]"),
]
FIGURE_CAPTION = re.compile(r"^图\s*\d+(?:[-.–—]\d+)?")
TABLE_CAPTION = re.compile(r"^表\s*\d+(?:[-.–—]\d+)?")
EQUATION_CAPTION = re.compile(r"^公式\s*\d+|^\(\s*\d+\s*\)\s*$")
TOC_STYLE = re.compile(r"(?i)^toc")
EMU_PER_INCH = 914400


def emu_to_inches(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return round(int(value) / EMU_PER_INCH, 4)
    except ValueError:
        return None


def node_text(node: Node) -> str:
    return "".join(child.data for child in node.childNodes if child.nodeType == Node.TEXT_NODE)


def paragraph_text(paragraph: Node) -> str:
    parts: list[str] = []

    def visit(node: Node) -> None:
        for child in node.childNodes:
            if child.nodeType != Node.ELEMENT_NODE:
                continue
            if child.nodeName == "w:p" and child is not paragraph:
                continue
            if child.nodeName == "w:t":
                parts.append(node_text(child))
            else:
                visit(child)

    visit(paragraph)
    return "".join(parts)


def ancestor(node: Node | None, tag: str) -> Node | None:
    current = node
    while current is not None:
        if getattr(current, "nodeName", None) == tag:
            return current
        current = getattr(current, "parentNode", None)
    return None


def heading_level(style_id: str | None, style_name: str, outline: int | None) -> int | None:
    if style_id in HEADING_STYLE_IDS:
        return HEADING_STYLE_IDS[style_id]
    match = re.match(r"(?i)^heading\s*(\d+)$", style_name or "")
    if match:
        return int(match.group(1))
    if outline is not None and 0 <= outline <= 8:
        return outline + 1
    return None


def classify_region_title(text: str) -> str | None:
    compact = re.sub(r"\s+", "", text).strip()
    spaced = text.strip()
    for region, pattern in REGION_PATTERNS:
        if pattern.match(spaced) or pattern.match(compact):
            return region
    return None


def looks_like(patterns: list[re.Pattern[str]], text: str) -> bool:
    value = text.strip()
    return any(pattern.search(value) for pattern in patterns)


def paragraph_blocks(document) -> list[dict[str, Any]]:
    bodies = document.getElementsByTagName("w:body")
    if not bodies:
        return []
    blocks: list[dict[str, Any]] = []
    section_index = 0
    paragraph_index = 0
    for block in element_children(bodies[0]):
        paragraphs: list[Node] = []
        if block.nodeName == "w:p":
            paragraphs.append(block)
        paragraphs.extend(list(block.getElementsByTagName("w:p")))
        seen: set[int] = set()
        for paragraph in paragraphs:
            identity = id(paragraph)
            if identity in seen:
                continue
            seen.add(identity)
            p_pr = direct_child(paragraph, "w:pPr")
            style_id = attr(direct_child(p_pr, "w:pStyle"), "w:val")
            props = paragraph_props(p_pr)
            r_pr = None
            runs = [child for child in paragraph.getElementsByTagName("w:r")]
            if runs:
                r_pr = direct_child(runs[0], "w:rPr")
            text = paragraph_text(paragraph)
            in_table = ancestor(paragraph.parentNode, "w:tbl") is not None
            in_text_box = ancestor(paragraph.parentNode, "w:txbxContent") is not None
            has_drawing = bool(paragraph.getElementsByTagName("w:drawing") or paragraph.getElementsByTagName("v:shape"))
            drawing_width = None
            drawing_height = None
            for extent in paragraph.getElementsByTagName("wp:extent"):
                drawing_width = emu_to_inches(extent.getAttribute("cx"))
                drawing_height = emu_to_inches(extent.getAttribute("cy"))
                break
            blocks.append(
                {
                    "index": paragraph_index,
                    "section": section_index,
                    "styleId": style_id,
                    "text": text,
                    "preview": re.sub(r"\s+", " ", text).strip()[:120],
                    "paragraph": props,
                    "run": run_props(r_pr),
                    "inTable": in_table,
                    "inTextBox": in_text_box,
                    "hasDrawing": has_drawing,
                    "hasTable": block.nodeName == "w:tbl" or in_table,
                    "drawingWidthIn": drawing_width,
                    "drawingHeightIn": drawing_height,
                }
            )
            paragraph_index += 1
        if block.nodeName == "w:p" and direct_child(direct_child(block, "w:pPr"), "w:sectPr") is not None:
            section_index += 1
        elif block.nodeName == "w:sectPr":
            section_index += 1
    return blocks


def enrich_style_names(blocks: list[dict[str, Any]], styles: dict[str, Any]) -> None:
    for item in blocks:
        style_id = item.get("styleId")
        catalog = styles.get(style_id or "") or {}
        item["styleName"] = catalog.get("name") or style_id or ""
        item["outlineLevel"] = catalog.get("outlineLevel")
        if item["outlineLevel"] is None:
            item["outlineLevel"] = item.get("paragraph", {}).get("outlineLevel")
        style_paragraph = catalog.get("paragraph") or {}
        if item.get("paragraph", {}).get("keepNext") is None and style_paragraph.get("keepNext") is not None:
            item.setdefault("paragraph", {})["keepNext"] = style_paragraph.get("keepNext")


def classify_block(item: dict[str, Any], region: str) -> tuple[str, str]:
    text = re.sub(r"\s+", " ", item.get("text") or "").strip()
    style_name = str(item.get("styleName") or "")
    style_id = str(item.get("styleId") or "")
    level = heading_level(style_id, style_name, item.get("outlineLevel"))

    if TOC_STYLE.search(style_name) or TOC_STYLE.search(style_id):
        return "toc_entry", "toc"
    if item.get("inTextBox") and region in {"cover", "front"}:
        return "cover_field", "cover"

    region_title = classify_region_title(text)
    if region_title and (level in {1, 2, None} or len(text) <= 20):
        return region_title, region_title

    if FIGURE_CAPTION.search(text):
        return "figure_caption", region
    if TABLE_CAPTION.search(text):
        return "table_caption", region
    if EQUATION_CAPTION.search(text):
        return "equation", region
    if item.get("hasDrawing") and not text:
        return "figure", region
    if item.get("inTable"):
        return "table_cell", region

    if level == 1 or looks_like(CHAPTER_PATTERNS, text) and level in {1, None} and len(text) < 80:
        if region_title:
            return region_title, region_title
        if looks_like(CHAPTER_PATTERNS, text) or level == 1:
            if classify_region_title(text):
                found = classify_region_title(text)
                return found or "chapter", found or "chapter"
            return "chapter", "chapter"
    if level == 2 or looks_like(SECTION_PATTERNS, text) and len(text) < 80:
        return "section", region if region != "front" else "chapter"
    if level == 3 or looks_like(SUBSECTION_PATTERNS, text) and len(text) < 80:
        return "subsection", region if region != "front" else "chapter"
    if not text:
        return "empty", region
    if region == "cover":
        return "cover", region
    if region.startswith("abstract"):
        return "abstract_body", region
    if region == "references":
        return "reference_item", region
    if region == "toc":
        return "toc_entry", region
    return "body", region if region != "front" else "body"


def build_tree(classified: list[dict[str, Any]], document_type: str) -> dict[str, Any]:
    root_children: list[dict[str, Any]] = []
    current_chapter: dict[str, Any] | None = None
    current_section: dict[str, Any] | None = None
    region_nodes: dict[str, dict[str, Any]] = {}

    def add_region(kind: str, node: dict[str, Any]) -> None:
        if kind in {"chapter", "section", "subsection", "body", "figure", "figure_caption", "table_caption", "table_cell", "equation", "empty"}:
            if current_section is not None and kind != "chapter":
                current_section.setdefault("children", []).append(node)
            elif current_chapter is not None and kind != "chapter":
                current_chapter.setdefault("children", []).append(node)
            else:
                root_children.append(node)
            return
        existing = region_nodes.get(kind)
        if existing is None:
            region_nodes[kind] = node
            root_children.append(node)
        else:
            existing.setdefault("children", []).append(
                {"type": node["type"], "paragraphIndex": node.get("paragraphIndex"), "title": node.get("title")}
            )

    chapter_number = 0
    for item in classified:
        kind = item["type"]
        node = {
            "type": kind,
            "paragraphIndex": item["index"],
            "section": item["section"],
            "styleId": item.get("styleId"),
            "styleName": item.get("styleName"),
            "title": item.get("preview") or None,
        }
        if kind == "chapter":
            chapter_number += 1
            node["number"] = chapter_number
            node["children"] = []
            current_chapter = node
            current_section = None
            root_children.append(node)
            continue
        if kind == "section":
            node["children"] = []
            current_section = node
            if current_chapter is not None:
                current_chapter.setdefault("children", []).append(node)
            else:
                root_children.append(node)
            continue
        add_region(kind, node)

    return {"type": document_type, "children": root_children}


def summarize(classified: list[dict[str, Any]], tree: dict[str, Any]) -> dict[str, Any]:
    counts: dict[str, int] = {}
    for item in classified:
        counts[item["type"]] = counts.get(item["type"], 0) + 1
    chapters = [child for child in tree.get("children", []) if child.get("type") == "chapter"]
    return {
        "paragraphs": len(classified),
        "counts": counts,
        "chapterCount": len(chapters),
        "hasCover": any(child.get("type") == "cover" for child in tree.get("children", [])),
        "hasAbstractZh": any(child.get("type") == "abstract_zh" for child in tree.get("children", [])),
        "hasAbstractEn": any(child.get("type") == "abstract_en" for child in tree.get("children", [])),
        "hasToc": any(child.get("type") in {"toc", "toc_entry"} for child in tree.get("children", []))
        or counts.get("toc_entry", 0) > 0,
        "hasReferences": any(child.get("type") == "references" for child in tree.get("children", [])),
        "hasAppendix": any(child.get("type") == "appendix" for child in tree.get("children", [])),
        "hasDeclaration": any(child.get("type") == "declaration" for child in tree.get("children", [])),
        "hasAcknowledgement": any(child.get("type") == "acknowledgement" for child in tree.get("children", [])),
        "hasListOfFigures": any(child.get("type") == "list_of_figures" for child in tree.get("children", [])),
        "hasListOfTables": any(child.get("type") == "list_of_tables" for child in tree.get("children", [])),
        "figureCaptions": counts.get("figure_caption", 0),
        "tableCaptions": counts.get("table_caption", 0),
        "equationCount": counts.get("equation", 0),
    }


def document_spec(profile: dict[str, Any] | None) -> dict[str, Any]:
    document = (profile or {}).get("document") or {}
    if not isinstance(document, dict):
        document = {}
    return {
        "type": document.get("type") or "thesis",
        "sections": document.get("sections") or [],
    }


def build_model(path: Path, profile: dict[str, Any] | None = None) -> dict[str, Any]:
    report = audit(path)
    with zipfile.ZipFile(path, "r") as archive:
        document = parse_xml(archive.read("word/document.xml"))
    blocks = paragraph_blocks(document)
    enrich_style_names(blocks, report.get("styles") or {})
    spec = document_spec(profile)
    region = "cover"
    classified: list[dict[str, Any]] = []
    seen_body = False
    for item in blocks:
        kind, region = classify_block(item, region)
        if kind == "chapter":
            seen_body = True
        if region == "cover" and seen_body:
            region = "chapter"
        item["type"] = kind
        item["region"] = region
        classified.append(item)
    tree = build_tree(classified, spec["type"])
    return {
        "path": str(path.resolve()),
        "documentType": spec["type"],
        "tree": tree,
        "summary": summarize(classified, tree),
        "blocks": [
            {
                "index": item["index"],
                "section": item["section"],
                "type": item["type"],
                "region": item["region"],
                "styleId": item.get("styleId"),
                "styleName": item.get("styleName"),
                "preview": item.get("preview"),
                "text": item.get("text"),
                "inTable": item.get("inTable"),
                "inTextBox": item.get("inTextBox"),
                "hasDrawing": item.get("hasDrawing"),
                "keepNext": (item.get("paragraph") or {}).get("keepNext"),
                "keepLines": (item.get("paragraph") or {}).get("keepLines"),
                "pageBreakBefore": (item.get("paragraph") or {}).get("pageBreakBefore"),
                "textLength": len((item.get("text") or "").strip()),
                "sizePt": (item.get("run") or {}).get("sizePt"),
                "lineSpacing": (item.get("paragraph") or {}).get("lineSpacing"),
                "drawingWidthIn": item.get("drawingWidthIn"),
                "drawingHeightIn": item.get("drawingHeightIn"),
            }
            for item in classified
        ],
        "audit": {
            "sections": report.get("sections"),
            "defaultRun": report.get("defaultRun"),
            "styles": report.get("styles"),
            "fields": report.get("fields"),
            "content": report.get("content"),
            "package": report.get("package"),
            "pageNumbers": [
                {"index": section.get("index"), "pageNumber": section.get("pageNumber"), "page": section.get("page")}
                for section in report.get("sections", [])
            ],
        },
    }


def load_profile(path: Path | None) -> dict[str, Any] | None:
    if path is None:
        return None
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Profile root must be an object")
    return data


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a semantic Document AST from a DOCX.")
    parser.add_argument("--input", required=True, type=Path, help="Source .docx")
    parser.add_argument("--profile", type=Path, help="Optional profile with document.spec")
    parser.add_argument("--output", type=Path, help="Write JSON model")
    args = parser.parse_args()
    try:
        if not args.input.is_file():
            raise FileNotFoundError(args.input)
        model = build_model(args.input, load_profile(args.profile))
        rendered = json.dumps(model, ensure_ascii=False, indent=2)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(rendered + "\n", encoding="utf-8")
            print(json.dumps({"ok": True, "output": str(args.output.resolve())}, ensure_ascii=False))
        else:
            print(rendered)
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
