#!/usr/bin/env python3
"""Read-only DOCX formatting inventory using only the Python standard library."""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any
from xml.dom import Node, minidom


REQUIRED_PARTS = {
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
}


def element_children(node: Node) -> list[Node]:
    return [child for child in node.childNodes if child.nodeType == Node.ELEMENT_NODE]


def direct_child(node: Node | None, tag: str) -> Node | None:
    if node is None:
        return None
    return next((child for child in element_children(node) if child.nodeName == tag), None)


def attr(node: Node | None, name: str) -> str | None:
    if node is None or not getattr(node, "hasAttribute", lambda _: False)(name):
        return None
    return node.getAttribute(name)


def parse_xml(data: bytes) -> minidom.Document:
    return minidom.parseString(data)


def twips_to_inches(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return round(int(value) / 1440, 4)
    except ValueError:
        return None


def half_points_to_points(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return int(value) / 2
    except ValueError:
        return None


def document_sections(document: minidom.Document) -> list[Node]:
    bodies = document.getElementsByTagName("w:body")
    if not bodies:
        return []
    sections: list[Node] = []
    for child in element_children(bodies[0]):
        if child.nodeName == "w:p":
            sect_pr = direct_child(direct_child(child, "w:pPr"), "w:sectPr")
            if sect_pr is not None:
                sections.append(sect_pr)
        elif child.nodeName == "w:sectPr":
            sections.append(child)
    return sections


def document_relationships(parts: dict[str, bytes]) -> dict[str, dict[str, str]]:
    part_name = "word/_rels/document.xml.rels"
    if part_name not in parts:
        return {}
    document = parse_xml(parts[part_name])
    out: dict[str, dict[str, str]] = {}
    for node in document.getElementsByTagName("Relationship"):
        relationship_id = attr(node, "Id")
        target = attr(node, "Target")
        if not relationship_id or not target:
            continue
        resolved = target if target.startswith("/") else posixpath.normpath(posixpath.join("word", target))
        out[relationship_id] = {
            "type": attr(node, "Type") or "",
            "target": resolved.lstrip("/"),
        }
    return out


def section_report(
    section: Node, index: int, relationships: dict[str, dict[str, str]]
) -> dict[str, Any]:
    pg_sz = direct_child(section, "w:pgSz")
    pg_mar = direct_child(section, "w:pgMar")
    pg_num = direct_child(section, "w:pgNumType")
    section_type = direct_child(section, "w:type")
    def references(tag: str) -> list[dict[str, Any]]:
        out = []
        for node in element_children(section):
            if node.nodeName != tag:
                continue
            relationship_id = attr(node, "r:id")
            relationship = relationships.get(relationship_id or "", {})
            out.append(
                {
                    "type": attr(node, "w:type") or "default",
                    "relationshipId": relationship_id,
                    "target": relationship.get("target"),
                }
            )
        return out

    header_references = references("w:headerReference")
    footer_references = references("w:footerReference")
    return {
        "index": index,
        "type": attr(section_type, "w:val") or "nextPage/default",
        "page": {
            "widthIn": twips_to_inches(attr(pg_sz, "w:w")),
            "heightIn": twips_to_inches(attr(pg_sz, "w:h")),
            "orientation": attr(pg_sz, "w:orient") or "portrait/default",
            "marginsIn": {
                "top": twips_to_inches(attr(pg_mar, "w:top")),
                "right": twips_to_inches(attr(pg_mar, "w:right")),
                "bottom": twips_to_inches(attr(pg_mar, "w:bottom")),
                "left": twips_to_inches(attr(pg_mar, "w:left")),
                "header": twips_to_inches(attr(pg_mar, "w:header")),
                "footer": twips_to_inches(attr(pg_mar, "w:footer")),
                "gutter": twips_to_inches(attr(pg_mar, "w:gutter")),
            },
        },
        "pageNumber": {
            "format": attr(pg_num, "w:fmt"),
            "start": int(attr(pg_num, "w:start")) if (attr(pg_num, "w:start") or "").isdigit() else None,
        },
        "differentFirstPage": direct_child(section, "w:titlePg") is not None,
        "headerReferences": header_references,
        "footerReferences": footer_references,
    }


def toggle_on(node: Node | None) -> bool | None:
    if node is None:
        return None
    value = attr(node, "w:val")
    if value is None:
        return True
    return str(value).lower() not in {"0", "false", "off"}


def paragraph_props(p_pr: Node | None) -> dict[str, Any]:
    if p_pr is None:
        return {}
    spacing = direct_child(p_pr, "w:spacing")
    indent = direct_child(p_pr, "w:ind")
    jc = direct_child(p_pr, "w:jc")
    outline = direct_child(p_pr, "w:outlineLvl")
    line = attr(spacing, "w:line")
    line_rule = attr(spacing, "w:lineRule") or "auto"
    line_spacing = None
    line_spacing_exact_pt = None
    if line and line.lstrip("-").isdigit():
        value = int(line)
        if line_rule == "auto":
            line_spacing = round(value / 240, 4)
        else:
            line_spacing_exact_pt = round(value / 20, 4)
    first_line_chars = attr(indent, "w:firstLineChars")
    return {
        "alignment": attr(jc, "w:val"),
        "lineSpacing": line_spacing,
        "lineSpacingExactPt": line_spacing_exact_pt,
        "lineRule": line_rule if line else None,
        "spaceBeforePt": round(int(attr(spacing, "w:before")) / 20, 4)
        if (attr(spacing, "w:before") or "").lstrip("-").isdigit()
        else None,
        "spaceAfterPt": round(int(attr(spacing, "w:after")) / 20, 4)
        if (attr(spacing, "w:after") or "").lstrip("-").isdigit()
        else None,
        "leftIn": twips_to_inches(attr(indent, "w:left")),
        "rightIn": twips_to_inches(attr(indent, "w:right")),
        "firstLineIn": twips_to_inches(attr(indent, "w:firstLine")),
        "hangingIn": twips_to_inches(attr(indent, "w:hanging")),
        "firstLineChars": int(first_line_chars) / 100 if (first_line_chars or "").isdigit() else None,
        "keepNext": toggle_on(direct_child(p_pr, "w:keepNext")),
        "keepLines": toggle_on(direct_child(p_pr, "w:keepLines")),
        "pageBreakBefore": toggle_on(direct_child(p_pr, "w:pageBreakBefore")),
        "widowControl": toggle_on(direct_child(p_pr, "w:widowControl")),
        "outlineLevel": int(attr(outline, "w:val")) if (attr(outline, "w:val") or "").isdigit() else None,
    }


def run_props(r_pr: Node | None) -> dict[str, Any]:
    if r_pr is None:
        return {}
    r_fonts = direct_child(r_pr, "w:rFonts")
    size = direct_child(r_pr, "w:sz")
    return {
        "latin": attr(r_fonts, "w:ascii") or attr(r_fonts, "w:hAnsi"),
        "eastAsia": attr(r_fonts, "w:eastAsia"),
        "complexScript": attr(r_fonts, "w:cs"),
        "sizePt": half_points_to_points(attr(size, "w:val")),
        "bold": toggle_on(direct_child(r_pr, "w:b")),
        "italic": toggle_on(direct_child(r_pr, "w:i")),
        "color": attr(direct_child(r_pr, "w:color"), "w:val"),
    }


def default_run_report(styles: minidom.Document | None) -> dict[str, Any] | None:
    if styles is None:
        return None
    roots = styles.getElementsByTagName("w:styles")
    if not roots:
        return None
    doc_defaults = direct_child(roots[0], "w:docDefaults")
    r_pr = direct_child(direct_child(doc_defaults, "w:rPrDefault"), "w:rPr")
    return run_props(r_pr) or None


def default_paragraph_report(styles: minidom.Document | None) -> dict[str, Any] | None:
    if styles is None:
        return None
    roots = styles.getElementsByTagName("w:styles")
    if not roots:
        return None
    doc_defaults = direct_child(roots[0], "w:docDefaults")
    p_pr = direct_child(direct_child(doc_defaults, "w:pPrDefault"), "w:pPr")
    props = paragraph_props(p_pr)
    return props or None


def style_catalog(styles: minidom.Document | None) -> dict[str, Any]:
    if styles is None:
        return {}
    out: dict[str, Any] = {}
    for style in styles.getElementsByTagName("w:style"):
        style_id = attr(style, "w:styleId")
        if not style_id:
            continue
        name = direct_child(style, "w:name")
        p_pr = direct_child(style, "w:pPr")
        r_pr = direct_child(style, "w:rPr")
        run = run_props(r_pr)
        paragraph = paragraph_props(p_pr)
        out[style_id] = {
            "name": attr(name, "w:val"),
            "type": attr(style, "w:type"),
            "basedOn": attr(direct_child(style, "w:basedOn"), "w:val"),
            "next": attr(direct_child(style, "w:next"), "w:val"),
            "fonts": {
                "latin": run.get("latin"),
                "eastAsia": run.get("eastAsia"),
                "complexScript": run.get("complexScript"),
            },
            "sizePt": run.get("sizePt"),
            "bold": bool(run.get("bold")),
            "numbered": direct_child(p_pr, "w:numPr") is not None,
            "outlineLevel": paragraph.get("outlineLevel"),
            "run": run,
            "paragraph": paragraph,
        }
    return out


def paragraph_inventory(document: minidom.Document) -> dict[str, Any]:
    style_usage: Counter[str] = Counter()
    paragraphs_with_direct_properties = 0
    direct_formatted_runs = 0
    paragraphs = list(document.getElementsByTagName("w:p"))
    runs = list(document.getElementsByTagName("w:r"))
    for paragraph in paragraphs:
        p_pr = direct_child(paragraph, "w:pPr")
        p_style = direct_child(p_pr, "w:pStyle")
        style_usage[attr(p_style, "w:val") or "(none)"] += 1
        if p_pr is not None:
            meaningful = [
                child
                for child in element_children(p_pr)
                if child.nodeName not in {"w:pStyle", "w:sectPr", "w:pPrChange"}
            ]
            if meaningful:
                paragraphs_with_direct_properties += 1
    for run in runs:
        if direct_child(run, "w:rPr") is not None:
            direct_formatted_runs += 1
    return {
        "paragraphs": len(paragraphs),
        "runs": len(runs),
        "styleUsage": dict(style_usage.most_common()),
        "paragraphsWithDirectProperties": paragraphs_with_direct_properties,
        "directFormattedRuns": direct_formatted_runs,
    }


def field_inventory(parts: dict[str, bytes]) -> dict[str, Any]:
    instructions: list[str] = []
    candidates = [
        name
        for name in parts
        if name == "word/document.xml"
        or re.fullmatch(r"word/(header|footer)\d+\.xml", name)
        or name in {"word/footnotes.xml", "word/endnotes.xml"}
    ]
    for name in candidates:
        try:
            dom = parse_xml(parts[name])
        except Exception:
            continue
        for node in dom.getElementsByTagName("w:instrText"):
            text = "".join(child.data for child in node.childNodes if child.nodeType == Node.TEXT_NODE)
            if text.strip():
                instructions.append(" ".join(text.split()))
        for node in dom.getElementsByTagName("w:fldSimple"):
            value = attr(node, "w:instr")
            if value:
                instructions.append(" ".join(value.split()))
    kinds: Counter[str] = Counter()
    for instruction in instructions:
        match = re.match(r"\s*([A-Za-z]+)", instruction)
        kinds[(match.group(1).upper() if match else "UNKNOWN")] += 1
    field_starts = 0
    for name in candidates:
        try:
            dom = parse_xml(parts[name])
        except Exception:
            continue
        field_starts += len(
            [
                node
                for node in dom.getElementsByTagName("w:fldChar")
                if attr(node, "w:fldCharType") == "begin"
            ]
        )
    return {
        "count": len(instructions),
        "fieldStarts": field_starts,
        "kinds": dict(kinds),
        "instructions": instructions,
    }


def text_metrics(document: minidom.Document) -> dict[str, int]:
    text = "".join(
        "".join(child.data for child in node.childNodes if child.nodeType == Node.TEXT_NODE)
        for node in document.getElementsByTagName("w:t")
    )
    return {
        "characters": len(text),
        "charactersNoWhitespace": len(re.sub(r"\s+", "", text)),
        "cjkCharacters": len(re.findall(r"[\u3400-\u4dbf\u4e00-\u9fff]", text)),
        "latinWords": len(re.findall(r"[A-Za-z]+(?:['’-][A-Za-z]+)*", text)),
    }


def bookmark_inventory(document: minidom.Document) -> dict[str, Any]:
    names = [
        attr(node, "w:name")
        for node in document.getElementsByTagName("w:bookmarkStart")
        if attr(node, "w:name")
    ]
    return {"count": len(names), "names": names}


def object_inventory(document: minidom.Document) -> dict[str, int]:
    return {
        "drawings": len(document.getElementsByTagName("w:drawing")),
        "anchoredDrawings": len(document.getElementsByTagName("wp:anchor")),
        "inlineDrawings": len(document.getElementsByTagName("wp:inline")),
        "vmlShapes": len(document.getElementsByTagName("v:shape")),
        "textBoxes": len(document.getElementsByTagName("w:txbxContent")),
        "contentControls": len(document.getElementsByTagName("w:sdt")),
    }


def marginal_inventory(parts: dict[str, bytes]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for name in sorted(parts):
        if not re.fullmatch(r"word/(header|footer)\d+\.xml", name):
            continue
        try:
            document = parse_xml(parts[name])
        except Exception:
            continue
        text = "".join(
            "".join(child.data for child in node.childNodes if child.nodeType == Node.TEXT_NODE)
            for node in document.getElementsByTagName("w:t")
        )
        instructions = []
        for node in document.getElementsByTagName("w:instrText"):
            value = "".join(child.data for child in node.childNodes if child.nodeType == Node.TEXT_NODE)
            if value.strip():
                instructions.append(" ".join(value.split()))
        out.append(
            {
                "part": name,
                "paragraphs": len(document.getElementsByTagName("w:p")),
                "text": text,
                "fields": instructions,
            }
        )
    return out


def numbering_inventory(parts: dict[str, bytes]) -> dict[str, Any]:
    if "word/numbering.xml" not in parts:
        return {"partPresent": False, "abstractDefinitions": 0, "instances": 0, "formats": {}}
    document = parse_xml(parts["word/numbering.xml"])
    formats: Counter[str] = Counter()
    for node in document.getElementsByTagName("w:numFmt"):
        formats[attr(node, "w:val") or "(missing)"] += 1
    return {
        "partPresent": True,
        "abstractDefinitions": len(document.getElementsByTagName("w:abstractNum")),
        "instances": len(document.getElementsByTagName("w:num")),
        "formats": dict(formats),
    }


def audit(path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(path, "r") as archive:
        infos = archive.infolist()
        names = [info.filename for info in infos]
        duplicate_names = sorted(name for name, count in Counter(names).items() if count > 1)
        corrupt_member = archive.testzip()
        parts = {info.filename: archive.read(info.filename) for info in infos}

    missing = sorted(REQUIRED_PARTS - parts.keys())
    if "word/document.xml" not in parts:
        raise ValueError("Missing word/document.xml")
    document = parse_xml(parts["word/document.xml"])
    relationships = document_relationships(parts)
    styles = parse_xml(parts["word/styles.xml"]) if "word/styles.xml" in parts else None
    settings = parse_xml(parts["word/settings.xml"]) if "word/settings.xml" in parts else None
    update_fields = None
    even_and_odd_headers = None
    if settings is not None:
        nodes = settings.getElementsByTagName("w:updateFields")
        if nodes:
            value = attr(nodes[0], "w:val")
            update_fields = value not in {"0", "false", "off"}
        nodes = settings.getElementsByTagName("w:evenAndOddHeaders")
        if nodes:
            value = attr(nodes[0], "w:val")
            even_and_odd_headers = value not in {"0", "false", "off"}

    sections = document_sections(document)
    return {
        "path": str(path.resolve()),
        "package": {
            "parts": len(parts),
            "missingRequiredParts": missing,
            "duplicateNames": duplicate_names,
            "corruptMember": corrupt_member,
        },
        "content": {
            **paragraph_inventory(document),
            **text_metrics(document),
            "tables": len(document.getElementsByTagName("w:tbl")),
            "images": len([name for name in parts if name.startswith("word/media/")]),
            "commentsPart": "word/comments.xml" in parts,
            "trackedInsertions": len(document.getElementsByTagName("w:ins")),
            "trackedDeletions": len(document.getElementsByTagName("w:del")),
            **object_inventory(document),
        },
        "defaultRun": default_run_report(styles),
        "defaultParagraph": default_paragraph_report(styles),
        "styles": style_catalog(styles),
        "sections": [
            section_report(section, index, relationships) for index, section in enumerate(sections)
        ],
        "fields": {**field_inventory(parts), "updateOnOpen": update_fields},
        "bookmarks": bookmark_inventory(document),
        "marginals": {
            "headers": len([name for name in parts if re.fullmatch(r"word/header\d+\.xml", name)]),
            "footers": len([name for name in parts if re.fullmatch(r"word/footer\d+\.xml", name)]),
            "evenAndOddHeaders": even_and_odd_headers,
            "parts": marginal_inventory(parts),
        },
        "numbering": numbering_inventory(parts),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit DOCX formatting without changing the file.")
    parser.add_argument("input", type=Path, help="Input .docx")
    parser.add_argument("--output", type=Path, help="Write JSON report to this path")
    args = parser.parse_args()

    try:
        if not args.input.is_file():
            raise FileNotFoundError(args.input)
        report = audit(args.input)
        rendered = json.dumps(report, ensure_ascii=False, indent=2)
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
