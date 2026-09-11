#!/usr/bin/env python3
"""Apply a JSON DOCX formatting profile with package-safe OOXML mutations."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any
from xml.dom import Node, minidom

from profile_schema import ensure_profile


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

SETTINGS_REL = f"{OFFICE_REL}/settings"
NUMBERING_REL = f"{OFFICE_REL}/numbering"
STYLES_REL = f"{OFFICE_REL}/styles"
HEADER_REL = f"{OFFICE_REL}/header"
FOOTER_REL = f"{OFFICE_REL}/footer"
SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
HEADER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FOOTER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"

PAGE_SIZES_TWIPS = {
    "a3": (16838, 23811),
    "a4": (11906, 16838),
    "a5": (8391, 11906),
    "letter": (12240, 15840),
    "legal": (12240, 20160),
    "tabloid": (15840, 24480),
}

NUMBER_FORMATS = {
    "decimal": "decimal",
    "lower-roman": "lowerRoman",
    "upper-roman": "upperRoman",
    "lower-alpha": "lowerLetter",
    "upper-alpha": "upperLetter",
    "chinese-counting": "chineseCounting",
    "chinese-counting-thousand": "chineseCountingThousand",
    "chinese-legal-simplified": "chineseLegalSimplified",
    "ideograph-traditional": "ideographTraditional",
    "ideograph-digital": "ideographDigital",
    "decimal-enclosed-circle": "decimalEnclosedCircle",
    "decimal-enclosed-fullstop": "decimalEnclosedFullstop",
    "decimal-full-width": "decimalFullWidth",
}

HEADER_FOOTER_TYPES = {"default", "first", "even"}
SECTION_TYPES = {"nextPage", "continuous", "evenPage", "oddPage", "nextColumn"}
ALLOWED_FIELD_KINDS = {
    "TOC", "PAGE", "NUMPAGES", "SECTIONPAGES", "SEQ", "REF", "PAGEREF",
    "STYLEREF", "CITATION", "BIBLIOGRAPHY", "NOTEREF", "HYPERLINK",
}

STYLE_ORDER = [
    "w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine",
    "w:hidden", "w:uiPriority", "w:semiHidden", "w:unhideWhenUsed", "w:qFormat",
    "w:locked", "w:personal", "w:personalCompose", "w:personalReply", "w:rsid",
    "w:pPr", "w:rPr",
]

PPR_ORDER = [
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
    "w:sectPr", "w:pPrChange",
]

RPR_ORDER = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
    "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss",
    "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color",
    "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight",
    "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl",
    "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath",
    "w:rPrChange",
]

SECTPR_ORDER = [
    "w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr", "w:type",
    "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType",
    "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection",
    "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange",
]

SETTINGS_AFTER_UPDATE = {
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars",
    "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
    "w:captions", "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary",
    "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
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


def insert_ordered(parent: Node, child: Node, order: list[str]) -> None:
    try:
        rank = order.index(child.nodeName)
    except ValueError:
        rank = len(order) - 0.5
    for existing in element_children(parent):
        try:
            existing_rank = order.index(existing.nodeName)
        except ValueError:
            existing_rank = len(order) - 0.5
        if existing_rank > rank:
            parent.insertBefore(child, existing)
            return
    parent.appendChild(child)


def ensure_direct(parent: Node, tag: str, order: list[str] | None = None) -> Node:
    existing = direct_child(parent, tag)
    if existing is not None:
        return existing
    child = parent.ownerDocument.createElement(tag)
    if order:
        insert_ordered(parent, child, order)
    else:
        parent.appendChild(child)
    return child


def set_val(node: Node, value: Any) -> None:
    node.setAttribute("w:val", str(value))


def set_toggle(parent: Node, tag: str, enabled: bool, order: list[str]) -> None:
    node = ensure_direct(parent, tag, order)
    set_val(node, "1" if enabled else "0")


def parse_xml(data: bytes) -> minidom.Document:
    return minidom.parseString(data)


def xml_bytes(document: minidom.Document) -> bytes:
    return document.toxml(encoding="UTF-8")


def inches_to_twips(value: Any, label: str) -> int:
    if not isinstance(value, (int, float)) or value < 0:
        raise ValueError(f"{label} must be a non-negative number")
    return round(float(value) * 1440)


def points_to_twips(value: Any, label: str) -> int:
    if not isinstance(value, (int, float)) or value < 0:
        raise ValueError(f"{label} must be a non-negative number")
    return round(float(value) * 20)


def points_to_half_points(value: Any, label: str) -> int:
    if not isinstance(value, (int, float)) or value <= 0:
        raise ValueError(f"{label} must be a positive number")
    return round(float(value) * 2)


def document_sections(document: minidom.Document) -> list[Node]:
    bodies = document.getElementsByTagName("w:body")
    if not bodies:
        raise ValueError("word/document.xml has no w:body")
    sections: list[Node] = []
    for child in element_children(bodies[0]):
        if child.nodeName == "w:p":
            sect_pr = direct_child(direct_child(child, "w:pPr"), "w:sectPr")
            if sect_pr is not None:
                sections.append(sect_pr)
        elif child.nodeName == "w:sectPr":
            sections.append(child)
    if not sections:
        raise ValueError("Document has no section properties")
    return sections


def requested_page_size(config: dict[str, Any]) -> tuple[int, int] | None:
    size_name = config.get("size")
    if size_name is not None:
        if not isinstance(size_name, str) or size_name.lower() not in PAGE_SIZES_TWIPS:
            raise ValueError(f"Unsupported page size: {size_name}")
        return PAGE_SIZES_TWIPS[size_name.lower()]
    if "widthIn" in config or "heightIn" in config:
        if "widthIn" not in config or "heightIn" not in config:
            raise ValueError("Custom page size needs both widthIn and heightIn")
        return (
            inches_to_twips(config["widthIn"], "page.widthIn"),
            inches_to_twips(config["heightIn"], "page.heightIn"),
        )
    return None


def configure_page_on_section(section: Node, config: dict[str, Any], label: str) -> None:
    orientation = config.get("orientation")
    if orientation not in {None, "portrait", "landscape"}:
        raise ValueError(f"{label}.orientation must be portrait or landscape")
    requested_size = requested_page_size(config)
    margins = config.get("marginsIn", {})
    if not isinstance(margins, dict):
        raise ValueError(f"{label}.marginsIn must be an object")

    pg_sz = direct_child(section, "w:pgSz")
    if requested_size is not None or orientation is not None:
        pg_sz = pg_sz or ensure_direct(section, "w:pgSz", SECTPR_ORDER)
        if requested_size is not None:
            width, height = requested_size
        else:
            try:
                width = int(pg_sz.getAttribute("w:w"))
                height = int(pg_sz.getAttribute("w:h"))
            except (TypeError, ValueError):
                width, height = PAGE_SIZES_TWIPS["a4"]
        short, long = sorted((width, height))
        effective_orientation = orientation or (
            "landscape" if pg_sz.getAttribute("w:orient") == "landscape" or width > height else "portrait"
        )
        if effective_orientation == "landscape":
            pg_sz.setAttribute("w:w", str(long))
            pg_sz.setAttribute("w:h", str(short))
            pg_sz.setAttribute("w:orient", "landscape")
        else:
            pg_sz.setAttribute("w:w", str(short))
            pg_sz.setAttribute("w:h", str(long))
            if pg_sz.hasAttribute("w:orient"):
                pg_sz.removeAttribute("w:orient")
    if margins:
        pg_mar = ensure_direct(section, "w:pgMar", SECTPR_ORDER)
        for edge in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
            if edge in margins:
                pg_mar.setAttribute(
                    f"w:{edge}", str(inches_to_twips(margins[edge], f"{label}.marginsIn.{edge}"))
                )


def apply_page(document: minidom.Document, config: dict[str, Any], summary: dict[str, Any]) -> None:
    if not config:
        return
    sections = document_sections(document)
    for section in sections:
        configure_page_on_section(section, config, "page")
    summary["sectionsPageConfigured"] = len(sections)


def apply_section_configs(document: minidom.Document, configs: list[Any], summary: dict[str, Any]) -> None:
    if not configs:
        return
    if not isinstance(configs, list):
        raise ValueError("sections must be an array")
    sections = document_sections(document)
    changed = 0
    for item in configs:
        if not isinstance(item, dict) or not isinstance(item.get("section"), int):
            raise ValueError("Each sections item needs integer section")
        index = item["section"]
        if index < 0 or index >= len(sections):
            raise ValueError(f"sections section {index} is out of range; document has {len(sections)} sections")
        section = sections[index]
        page_config = item.get("page", {})
        if page_config:
            if not isinstance(page_config, dict):
                raise ValueError(f"sections[{index}].page must be an object")
            configure_page_on_section(section, page_config, f"sections[{index}].page")
        if "type" in item:
            section_type = item["type"]
            if section_type not in SECTION_TYPES:
                raise ValueError(f"Unsupported section type: {section_type}")
            set_val(ensure_direct(section, "w:type", SECTPR_ORDER), section_type)
        if "differentFirstPage" in item:
            set_toggle(section, "w:titlePg", bool(item["differentFirstPage"]), SECTPR_ORDER)
        changed += 1
    summary["sectionsConfigured"] = changed


def ensure_styles_root(parts: dict[str, bytes]) -> tuple[minidom.Document, Node]:
    if "word/styles.xml" in parts:
        document = parse_xml(parts["word/styles.xml"])
        roots = document.getElementsByTagName("w:styles")
        if not roots:
            raise ValueError("word/styles.xml has no w:styles root")
        return document, roots[0]
    document = minidom.parseString(f'<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="{W_NS}"/>')
    return document, document.documentElement


def apply_fonts(r_pr: Node, config: dict[str, Any]) -> None:
    if not config:
        return
    if any(key in config for key in ("latinFont", "eastAsiaFont", "complexScriptFont")):
        fonts = ensure_direct(r_pr, "w:rFonts", RPR_ORDER)
        if "latinFont" in config:
            fonts.setAttribute("w:ascii", str(config["latinFont"]))
            fonts.setAttribute("w:hAnsi", str(config["latinFont"]))
            for key in ("w:asciiTheme", "w:hAnsiTheme"):
                if fonts.hasAttribute(key):
                    fonts.removeAttribute(key)
        if "eastAsiaFont" in config:
            fonts.setAttribute("w:eastAsia", str(config["eastAsiaFont"]))
            if fonts.hasAttribute("w:eastAsiaTheme"):
                fonts.removeAttribute("w:eastAsiaTheme")
        if "complexScriptFont" in config:
            fonts.setAttribute("w:cs", str(config["complexScriptFont"]))
            if fonts.hasAttribute("w:cstheme"):
                fonts.removeAttribute("w:cstheme")
    if "sizePt" in config:
        half_points = points_to_half_points(config["sizePt"], "sizePt")
        set_val(ensure_direct(r_pr, "w:sz", RPR_ORDER), half_points)
        set_val(ensure_direct(r_pr, "w:szCs", RPR_ORDER), half_points)
    for key, tag in (("bold", "w:b"), ("italic", "w:i")):
        if key in config:
            set_toggle(r_pr, tag, bool(config[key]), RPR_ORDER)
    if "color" in config:
        color = str(config["color"]).upper().lstrip("#")
        if not re.fullmatch(r"[0-9A-F]{6}", color):
            raise ValueError(f"Invalid color: {config['color']}")
        set_val(ensure_direct(r_pr, "w:color", RPR_ORDER), color)


def apply_default_run(styles_root: Node, config: dict[str, Any]) -> None:
    if not config:
        return
    doc_defaults = ensure_direct(styles_root, "w:docDefaults", ["w:docDefaults", "w:latentStyles", "w:style"])
    r_pr_default = ensure_direct(doc_defaults, "w:rPrDefault", ["w:rPrDefault", "w:pPrDefault"])
    r_pr = ensure_direct(r_pr_default, "w:rPr")
    apply_fonts(r_pr, config)


def find_style(styles_root: Node, style_id: str) -> Node | None:
    for style in element_children(styles_root):
        if style.nodeName == "w:style" and style.getAttribute("w:styleId") == style_id:
            return style
    return None


def ensure_style(styles_root: Node, style_id: str, config: dict[str, Any]) -> Node:
    style = find_style(styles_root, style_id)
    if style is not None:
        return style
    style = styles_root.ownerDocument.createElement("w:style")
    style.setAttribute("w:type", "paragraph")
    style.setAttribute("w:styleId", style_id)
    name = ensure_direct(style, "w:name", STYLE_ORDER)
    name.setAttribute("w:val", str(config.get("name", style_id)))
    if style_id != "Normal":
        based_on = ensure_direct(style, "w:basedOn", STYLE_ORDER)
        based_on.setAttribute("w:val", str(config.get("basedOn", "Normal")))
        next_style = ensure_direct(style, "w:next", STYLE_ORDER)
        next_style.setAttribute("w:val", str(config.get("next", "Normal")))
    if style_id.startswith("Heading"):
        ensure_direct(style, "w:qFormat", STYLE_ORDER)
    styles_root.appendChild(style)
    return style


def apply_paragraph_format(p_pr: Node, config: dict[str, Any], size_pt: float | None) -> None:
    if not config:
        return
    alignment_map = {"left": "left", "center": "center", "right": "right", "justify": "both"}
    if "alignment" in config:
        alignment = config["alignment"]
        if alignment not in alignment_map:
            raise ValueError(f"Invalid alignment: {alignment}")
        set_val(ensure_direct(p_pr, "w:jc", PPR_ORDER), alignment_map[alignment])

    spacing_keys = {"spaceBeforePt", "spaceAfterPt", "lineSpacing", "lineSpacingExactPt"}
    if any(key in config for key in spacing_keys):
        if "lineSpacing" in config and "lineSpacingExactPt" in config:
            raise ValueError("Use only one of lineSpacing and lineSpacingExactPt")
        spacing = ensure_direct(p_pr, "w:spacing", PPR_ORDER)
        if "spaceBeforePt" in config:
            spacing.setAttribute("w:before", str(points_to_twips(config["spaceBeforePt"], "spaceBeforePt")))
        if "spaceAfterPt" in config:
            spacing.setAttribute("w:after", str(points_to_twips(config["spaceAfterPt"], "spaceAfterPt")))
        if "lineSpacing" in config:
            multiplier = config["lineSpacing"]
            if not isinstance(multiplier, (int, float)) or multiplier <= 0:
                raise ValueError("lineSpacing must be a positive number")
            spacing.setAttribute("w:line", str(round(float(multiplier) * 240)))
            spacing.setAttribute("w:lineRule", "auto")
        if "lineSpacingExactPt" in config:
            spacing.setAttribute("w:line", str(points_to_twips(config["lineSpacingExactPt"], "lineSpacingExactPt")))
            spacing.setAttribute("w:lineRule", "exact")

    indent_keys = {"leftIn", "rightIn", "firstLineIn", "hangingIn", "firstLineChars"}
    if any(key in config for key in indent_keys):
        if "firstLineIn" in config and "firstLineChars" in config:
            raise ValueError("Use only one of firstLineIn and firstLineChars")
        indent = ensure_direct(p_pr, "w:ind", PPR_ORDER)
        mapping = {"leftIn": "left", "rightIn": "right", "firstLineIn": "firstLine", "hangingIn": "hanging"}
        for key, attribute in mapping.items():
            if key in config:
                indent.setAttribute(f"w:{attribute}", str(inches_to_twips(config[key], key)))
        if "firstLineChars" in config:
            chars = config["firstLineChars"]
            if not isinstance(chars, (int, float)) or chars < 0:
                raise ValueError("firstLineChars must be non-negative")
            effective_size = size_pt or 12
            indent.setAttribute("w:firstLine", str(round(float(chars) * effective_size * 20)))

    for key, tag in (
        ("keepNext", "w:keepNext"), ("keepLines", "w:keepLines"),
        ("pageBreakBefore", "w:pageBreakBefore"), ("widowControl", "w:widowControl"),
    ):
        if key in config:
            set_toggle(p_pr, tag, bool(config[key]), PPR_ORDER)
    if "outlineLevel" in config:
        level = config["outlineLevel"]
        if not isinstance(level, int) or not 0 <= level <= 8:
            raise ValueError("outlineLevel must be an integer from 0 to 8")
        set_val(ensure_direct(p_pr, "w:outlineLvl", PPR_ORDER), level)


def apply_styles(styles_root: Node, styles_config: dict[str, Any], default_run: dict[str, Any]) -> int:
    if not isinstance(styles_config, dict):
        raise ValueError("styles must be an object")
    changed = 0
    for style_id, config in styles_config.items():
        if not isinstance(config, dict):
            raise ValueError(f"styles.{style_id} must be an object")
        style = ensure_style(styles_root, style_id, config)
        for key, tag in (("name", "w:name"), ("basedOn", "w:basedOn"), ("next", "w:next")):
            if key in config:
                node = ensure_direct(style, tag, STYLE_ORDER)
                node.setAttribute("w:val", str(config[key]))
        run_config = config.get("run", {})
        if run_config:
            r_pr = ensure_direct(style, "w:rPr", STYLE_ORDER)
            apply_fonts(r_pr, run_config)
        paragraph_config = config.get("paragraph", {})
        if paragraph_config:
            p_pr = ensure_direct(style, "w:pPr", STYLE_ORDER)
            size_pt = run_config.get("sizePt", default_run.get("sizePt"))
            apply_paragraph_format(p_pr, paragraph_config, float(size_pt) if size_pt else None)
        changed += 1
    return changed


def apply_theme(parts: dict[str, bytes], config: dict[str, Any], summary: dict[str, Any]) -> None:
    if not config or config.get("updateTheme", True) is False:
        return
    part_name = "word/theme/theme1.xml"
    if part_name not in parts:
        summary["themeUpdated"] = False
        return
    document = parse_xml(parts[part_name])
    changed = False
    for collection_name in ("a:majorFont", "a:minorFont"):
        collections = document.getElementsByTagName(collection_name)
        for collection in collections:
            for config_key, tag in (
                ("latinFont", "a:latin"), ("eastAsiaFont", "a:ea"), ("complexScriptFont", "a:cs")
            ):
                if config_key not in config:
                    continue
                node = ensure_direct(collection, tag, ["a:latin", "a:ea", "a:cs", "a:font"])
                node.setAttribute("typeface", str(config[config_key]))
                if node.hasAttribute("panose"):
                    node.removeAttribute("panose")
                changed = True
    if changed:
        parts[part_name] = xml_bytes(document)
    summary["themeUpdated"] = changed


def apply_page_numbers(document: minidom.Document, configs: list[Any], summary: dict[str, Any]) -> None:
    if not configs:
        return
    if not isinstance(configs, list):
        raise ValueError("pageNumbers must be an array")
    sections = document_sections(document)
    changed = 0
    for item in configs:
        if not isinstance(item, dict) or not isinstance(item.get("section"), int):
            raise ValueError("Each pageNumbers item needs integer section")
        section_index = item["section"]
        if section_index < 0 or section_index >= len(sections):
            raise ValueError(f"pageNumbers section {section_index} is out of range; document has {len(sections)} sections")
        node = ensure_direct(sections[section_index], "w:pgNumType", SECTPR_ORDER)
        if "format" in item:
            friendly = item["format"]
            if friendly not in NUMBER_FORMATS:
                raise ValueError(f"Unsupported page number format: {friendly}")
            node.setAttribute("w:fmt", NUMBER_FORMATS[friendly])
        if "start" in item:
            start = item["start"]
            if not isinstance(start, int) or start < 1:
                raise ValueError("page number start must be a positive integer")
            node.setAttribute("w:start", str(start))
        changed += 1
    summary["pageNumberSectionsConfigured"] = changed


def ensure_content_type(parts: dict[str, bytes], part_name: str, content_type: str) -> None:
    document = parse_xml(parts["[Content_Types].xml"])
    root = document.documentElement
    normalized = "/" + part_name.lstrip("/")
    for node in document.getElementsByTagName("Override"):
        if node.getAttribute("PartName") == normalized:
            node.setAttribute("ContentType", content_type)
            parts["[Content_Types].xml"] = xml_bytes(document)
            return
    node = document.createElement("Override")
    node.setAttribute("PartName", normalized)
    node.setAttribute("ContentType", content_type)
    root.appendChild(node)
    parts["[Content_Types].xml"] = xml_bytes(document)


def ensure_relationship(parts: dict[str, bytes], rel_type: str, target: str) -> str:
    part_name = "word/_rels/document.xml.rels"
    if part_name in parts:
        document = parse_xml(parts[part_name])
    else:
        document = minidom.parseString(
            f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{REL_NS}"/>'
        )
    relationships = document.getElementsByTagName("Relationship")
    for node in relationships:
        if node.getAttribute("Type") == rel_type and node.getAttribute("Target") == target:
            parts[part_name] = xml_bytes(document)
            return node.getAttribute("Id")
    ids = []
    for node in relationships:
        match = re.fullmatch(r"rId(\d+)", node.getAttribute("Id"))
        if match:
            ids.append(int(match.group(1)))
    node = document.createElement("Relationship")
    relationship_id = f"rId{max(ids, default=0) + 1}"
    node.setAttribute("Id", relationship_id)
    node.setAttribute("Type", rel_type)
    node.setAttribute("Target", target)
    document.documentElement.appendChild(node)
    parts[part_name] = xml_bytes(document)
    return relationship_id


def ensure_settings(parts: dict[str, bytes]) -> minidom.Document:
    part_name = "word/settings.xml"
    document = (
        parse_xml(parts[part_name])
        if part_name in parts
        else minidom.parseString(
            f'<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="{W_NS}"/>'
        )
    )
    ensure_relationship(parts, SETTINGS_REL, "settings.xml")
    ensure_content_type(parts, part_name, SETTINGS_CT)
    return document


def apply_update_fields(parts: dict[str, bytes], config: dict[str, Any], summary: dict[str, Any]) -> None:
    if "updateOnOpen" not in config:
        return
    document = ensure_settings(parts)
    root = document.documentElement
    node = direct_child(root, "w:updateFields")
    if node is None:
        node = document.createElement("w:updateFields")
        inserted = False
        for existing in element_children(root):
            if existing.nodeName in SETTINGS_AFTER_UPDATE:
                root.insertBefore(node, existing)
                inserted = True
                break
        if not inserted:
            root.appendChild(node)
    node.setAttribute("w:val", "true" if bool(config["updateOnOpen"]) else "false")
    parts["word/settings.xml"] = xml_bytes(document)
    summary["updateFields"] = bool(config["updateOnOpen"])


def next_numeric_attribute(nodes: list[Node], attribute: str) -> int:
    values = []
    for node in nodes:
        value = node.getAttribute(attribute)
        if value.isdigit():
            values.append(int(value))
    return max(values, default=-1) + 1


def apply_heading_numbering(
    parts: dict[str, bytes], styles_root: Node, config: dict[str, Any], summary: dict[str, Any]
) -> None:
    if not config or config.get("enabled") is not True:
        return
    styles = config.get("styles")
    patterns = config.get("patterns")
    formats = config.get("formats")
    starts = config.get("starts")
    if not all(isinstance(value, list) for value in (styles, patterns, formats, starts)):
        raise ValueError("headingNumbering needs styles, patterns, formats, and starts arrays")
    length = len(styles)
    if length < 1 or length > 9 or any(len(value) != length for value in (patterns, formats, starts)):
        raise ValueError("headingNumbering arrays must have the same length from 1 to 9")
    suffix = config.get("suffix", "space")
    if suffix not in {"space", "tab", "nothing"}:
        raise ValueError("headingNumbering.suffix must be space, tab, or nothing")

    part_name = "word/numbering.xml"
    if part_name in parts:
        numbering = parse_xml(parts[part_name])
    else:
        numbering = minidom.parseString(f'<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="{W_NS}"/>')
    ensure_relationship(parts, NUMBERING_REL, "numbering.xml")
    ensure_content_type(parts, part_name, NUMBERING_CT)
    root = numbering.documentElement
    marker = "5354464D"
    abstract = next(
        (
            node
            for node in element_children(root)
            if node.nodeName == "w:abstractNum"
            and direct_child(node, "w:nsid") is not None
            and direct_child(node, "w:nsid").getAttribute("w:val") == marker
        ),
        None,
    )
    abstract_is_new = abstract is None
    if abstract is None:
        abstract_id = next_numeric_attribute(
            list(numbering.getElementsByTagName("w:abstractNum")), "w:abstractNumId"
        )
        abstract = numbering.createElement("w:abstractNum")
        abstract.setAttribute("w:abstractNumId", str(abstract_id))
    else:
        abstract_id = int(abstract.getAttribute("w:abstractNumId"))
        while abstract.firstChild is not None:
            abstract.removeChild(abstract.firstChild)

    num = next(
        (
            node
            for node in element_children(root)
            if node.nodeName == "w:num"
            and direct_child(node, "w:abstractNumId") is not None
            and direct_child(node, "w:abstractNumId").getAttribute("w:val") == str(abstract_id)
        ),
        None,
    )
    num_is_new = num is None
    if num is None:
        num_id = max(1, next_numeric_attribute(list(numbering.getElementsByTagName("w:num")), "w:numId"))
        num = numbering.createElement("w:num")
        num.setAttribute("w:numId", str(num_id))
    else:
        num_id = int(num.getAttribute("w:numId"))
        while num.firstChild is not None:
            num.removeChild(num.firstChild)

    nsid = numbering.createElement("w:nsid")
    nsid.setAttribute("w:val", marker)
    abstract.appendChild(nsid)
    multi = numbering.createElement("w:multiLevelType")
    multi.setAttribute("w:val", "multilevel")
    abstract.appendChild(multi)

    for level, (style_id, pattern, friendly_format, start) in enumerate(zip(styles, patterns, formats, starts)):
        if friendly_format not in NUMBER_FORMATS:
            raise ValueError(f"Unsupported heading number format: {friendly_format}")
        if not isinstance(start, int) or start < 1:
            raise ValueError("headingNumbering.starts values must be positive integers")
        if not isinstance(pattern, str) or f"%{level + 1}" not in pattern:
            raise ValueError(f"headingNumbering pattern for level {level + 1} must include %{level + 1}")
        level_node = numbering.createElement("w:lvl")
        level_node.setAttribute("w:ilvl", str(level))
        for tag, value in (
            ("w:start", start), ("w:numFmt", NUMBER_FORMATS[friendly_format]),
            ("w:pStyle", style_id), ("w:suff", suffix), ("w:lvlText", pattern), ("w:lvlJc", "left"),
        ):
            node = numbering.createElement(tag)
            node.setAttribute("w:val", str(value))
            level_node.appendChild(node)
        abstract.appendChild(level_node)

        style = ensure_style(styles_root, str(style_id), {"name": str(style_id)})
        p_pr = ensure_direct(style, "w:pPr", STYLE_ORDER)
        num_pr = ensure_direct(p_pr, "w:numPr", PPR_ORDER)
        ilvl = ensure_direct(num_pr, "w:ilvl", ["w:ilvl", "w:numId", "w:numberingChange", "w:ins"])
        set_val(ilvl, level)
        num_id_node = ensure_direct(num_pr, "w:numId", ["w:ilvl", "w:numId", "w:numberingChange", "w:ins"])
        set_val(num_id_node, num_id)

    if abstract_is_new:
        first_num = direct_child(root, "w:num")
        if first_num is not None:
            root.insertBefore(abstract, first_num)
        else:
            root.appendChild(abstract)
    abstract_ref = numbering.createElement("w:abstractNumId")
    abstract_ref.setAttribute("w:val", str(abstract_id))
    num.appendChild(abstract_ref)
    if num_is_new:
        num_id_mac = direct_child(root, "w:numIdMacAtCleanup")
        if num_id_mac is not None:
            root.insertBefore(num, num_id_mac)
        else:
            root.appendChild(num)
    parts[part_name] = xml_bytes(numbering)
    summary["headingNumbering"] = {"levels": length, "abstractNumId": abstract_id, "numId": num_id}


def text_value(node: Node) -> str:
    return "".join(child.data for child in node.childNodes if child.nodeType == Node.TEXT_NODE)


def set_text_value(node: Node, value: str) -> None:
    while node.firstChild is not None:
        node.removeChild(node.firstChild)
    node.appendChild(node.ownerDocument.createTextNode(value))
    if value.startswith((" ", "\t", "\n")) or value.endswith((" ", "\t", "\n")):
        node.setAttribute("xml:space", "preserve")
    elif node.hasAttribute("xml:space"):
        node.removeAttribute("xml:space")


def ancestor(node: Node | None, tag: str) -> Node | None:
    current = node
    while current is not None:
        if current.nodeType == Node.ELEMENT_NODE and current.nodeName == tag:
            return current
        current = current.parentNode
    return None


def paragraph_text_nodes(paragraph: Node) -> list[Node]:
    out: list[Node] = []

    def visit(node: Node) -> None:
        for child in node.childNodes:
            if child.nodeType != Node.ELEMENT_NODE:
                continue
            if child.nodeName == "w:p" and child is not paragraph:
                continue
            if child.nodeName == "w:t":
                out.append(child)
            else:
                visit(child)

    visit(paragraph)
    return out


def paragraph_runs(paragraph: Node) -> list[Node]:
    out: list[Node] = []

    def visit(node: Node) -> None:
        for child in node.childNodes:
            if child.nodeType != Node.ELEMENT_NODE:
                continue
            if child.nodeName == "w:p" and child is not paragraph:
                continue
            if child.nodeName == "w:r":
                out.append(child)
            else:
                visit(child)

    visit(paragraph)
    return out


def paragraph_text(paragraph: Node, normalize: bool = False) -> str:
    value = "".join(text_value(node) for node in paragraph_text_nodes(paragraph))
    return " ".join(value.split()) if normalize else value



def style_id_to_name(parts: dict[str, bytes]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    raw = parts.get("word/styles.xml")
    if not raw:
        return mapping
    document = parse_xml(raw)
    for style in document.getElementsByTagName("w:style"):
        style_id = style.getAttribute("w:styleId")
        name_node = direct_child(style, "w:name")
        mapping[style_id] = name_node.getAttribute("w:val") if name_node is not None else style_id
    return mapping


def paragraph_infos(document: minidom.Document, style_names: dict[str, str] | None = None) -> list[dict[str, Any]]:
    bodies = document.getElementsByTagName("w:body")
    if not bodies:
        return []
    infos: list[dict[str, Any]] = []
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
            style = attr(direct_child(p_pr, "w:pStyle"), "w:val")
            style_names = style_names or {}
            infos.append(
                {
                    "paragraph": paragraph,
                    "index": paragraph_index,
                    "section": section_index,
                    "style": style,
                    "styleName": style_names.get(style or "", style or ""),
                    "inTable": ancestor(paragraph.parentNode, "w:tbl") is not None,
                    "inTextBox": ancestor(paragraph.parentNode, "w:txbxContent") is not None,
                }
            )
            paragraph_index += 1
        if block.nodeName == "w:p" and direct_child(direct_child(block, "w:pPr"), "w:sectPr") is not None:
            section_index += 1
    return infos


def paragraph_matches(info: dict[str, Any], match: dict[str, Any]) -> bool:
    if not isinstance(match, dict):
        raise ValueError("paragraph match must be an object")
    paragraph = info["paragraph"]
    raw_text = paragraph_text(paragraph)
    normalized = " ".join(raw_text.split())
    text = normalized if match.get("normalizeWhitespace", True) else raw_text
    style_name = str(info.get("styleName") or "")
    style_id = str(info.get("style") or "")
    checks = (
        ("textEquals", lambda value: text == str(value)),
        ("textContains", lambda value: str(value) in text),
        ("startsWith", lambda value: text.startswith(str(value))),
        ("textRegex", lambda value: re.search(str(value), text) is not None),
        ("textNotRegex", lambda value: re.search(str(value), text) is None),
        ("currentStyle", lambda value: style_id == str(value)),
        ("currentStyleNotIn", lambda value: style_id not in set(value if isinstance(value, list) else [value])),
        ("styleNameRegex", lambda value: re.search(str(value), style_name, flags=re.IGNORECASE) is not None),
        ("styleNameNotRegex", lambda value: re.search(str(value), style_name, flags=re.IGNORECASE) is None),
        ("section", lambda value: info.get("section") == value),
        ("paragraphIndex", lambda value: info.get("index") == value),
        ("inTable", lambda value: info.get("inTable") is bool(value)),
        ("inTextBox", lambda value: info.get("inTextBox") is bool(value)),
    )
    for key, predicate in checks:
        if key in match and not predicate(match[key]):
            return False
    return True


def replace_span_in_paragraph(paragraph: Node, start: int, end: int, replacement: str) -> None:
    nodes = paragraph_text_nodes(paragraph)
    ranges: list[tuple[Node, int, int]] = []
    offset = 0
    for node in nodes:
        value = text_value(node)
        ranges.append((node, offset, offset + len(value)))
        offset += len(value)
    affected = [(node, left, right) for node, left, right in ranges if left < end and right > start]
    if not affected:
        raise ValueError("Replacement span does not map to paragraph text nodes")
    first, first_left, _ = affected[0]
    last, last_left, _ = affected[-1]
    first_value = text_value(first)
    last_value = text_value(last)
    prefix = first_value[: start - first_left]
    suffix = last_value[end - last_left :]
    if first is last:
        set_text_value(first, prefix + replacement + suffix)
        return
    set_text_value(first, prefix + replacement)
    for node, _, _ in affected[1:-1]:
        set_text_value(node, "")
    set_text_value(last, suffix)


def replace_text_in_paragraph(
    paragraph: Node, pattern: str, replacement: str, regex: bool, replace_all: bool
) -> int:
    original = paragraph_text(paragraph)
    if not original:
        return 0
    if regex:
        matches = list(re.finditer(pattern, original))
        if not replace_all:
            matches = matches[:1]
        spans = [(match.start(), match.end(), match.expand(replacement)) for match in matches]
    else:
        spans = []
        offset = 0
        while pattern:
            index = original.find(pattern, offset)
            if index < 0:
                break
            spans.append((index, index + len(pattern), replacement))
            if not replace_all:
                break
            offset = index + len(pattern)
    for start, end, value in reversed(spans):
        replace_span_in_paragraph(paragraph, start, end, value)
    return len(spans)


def apply_paragraph_rules(
    document: minidom.Document, rules: list[Any], default_run: dict[str, Any], summary: dict[str, Any], style_names: dict[str, str] | None = None
) -> None:
    if not rules:
        return
    if not isinstance(rules, list):
        raise ValueError("paragraphRules must be an array")
    infos = paragraph_infos(document, style_names=style_names)
    report: list[dict[str, Any]] = []
    for rule_index, rule in enumerate(rules):
        if not isinstance(rule, dict):
            raise ValueError("Each paragraphRules item must be an object")
        matched = [info for info in infos if paragraph_matches(info, rule.get("match", {}))]
        if rule.get("required") and not matched:
            raise ValueError(f"paragraphRules[{rule_index}] required match was not found")
        max_matches = rule.get("maxMatches")
        if max_matches is not None:
            if not isinstance(max_matches, int) or max_matches < 1:
                raise ValueError("paragraphRules.maxMatches must be a positive integer")
            matched = matched[:max_matches]
        changed = 0
        for info in matched:
            paragraph = info["paragraph"]
            p_pr = direct_child(paragraph, "w:pPr")
            if rule.get("remove"):
                if direct_child(p_pr, "w:sectPr") is not None:
                    raise ValueError("Refusing to remove a paragraph that owns section properties")
                paragraph.parentNode.removeChild(paragraph)
                changed += 1
                continue
            p_pr = p_pr or ensure_direct(paragraph, "w:pPr", ["w:pPr", "w:r", "w:hyperlink"])
            if rule.get("clearParagraphFormatting"):
                for child in list(element_children(p_pr)):
                    if child.nodeName not in {"w:pStyle", "w:sectPr", "w:pPrChange"}:
                        p_pr.removeChild(child)
            if "style" in rule:
                set_val(ensure_direct(p_pr, "w:pStyle", PPR_ORDER), str(rule["style"]))
            paragraph_config = rule.get("paragraph", {})
            if paragraph_config:
                if not isinstance(paragraph_config, dict):
                    raise ValueError("paragraphRules.paragraph must be an object")
                size = rule.get("run", {}).get("sizePt", default_run.get("sizePt"))
                apply_paragraph_format(p_pr, paragraph_config, float(size) if size else None)
            if rule.get("clearRunFormatting"):
                for run in paragraph_runs(paragraph):
                    r_pr = direct_child(run, "w:rPr")
                    if r_pr is not None:
                        run.removeChild(r_pr)
            run_config = rule.get("run", {})
            if run_config:
                if not isinstance(run_config, dict):
                    raise ValueError("paragraphRules.run must be an object")
                for run in paragraph_runs(paragraph):
                    apply_fonts(ensure_direct(run, "w:rPr", RPR_ORDER), run_config)
            text_change = rule.get("textRegexReplace")
            if text_change:
                if not isinstance(text_change, dict) or "pattern" not in text_change:
                    raise ValueError("textRegexReplace needs pattern and optional replacement")
                replace_text_in_paragraph(
                    paragraph,
                    str(text_change["pattern"]),
                    str(text_change.get("replacement", "")),
                    True,
                    bool(text_change.get("all", False)),
                )
            changed += 1
        report.append({"rule": rule_index, "matched": len(matched), "changed": changed})
    summary["paragraphRules"] = report


def selected_part_names(parts: dict[str, bytes], selectors: list[str] | None) -> list[str]:
    selectors = selectors or ["document", "headers", "footers"]
    if not isinstance(selectors, list):
        raise ValueError("parts selector must be an array")
    names: list[str] = []
    for selector in selectors:
        if selector == "document":
            candidates = ["word/document.xml"]
        elif selector == "headers":
            candidates = sorted(name for name in parts if re.fullmatch(r"word/header\d+\.xml", name))
        elif selector == "footers":
            candidates = sorted(name for name in parts if re.fullmatch(r"word/footer\d+\.xml", name))
        elif selector == "notes":
            candidates = [name for name in ("word/footnotes.xml", "word/endnotes.xml") if name in parts]
        elif selector == "all":
            candidates = selected_part_names(parts, ["document", "headers", "footers", "notes"])
        elif selector in parts and selector.endswith(".xml"):
            candidates = [selector]
        else:
            raise ValueError(f"Unsupported DOCX part selector: {selector}")
        for name in candidates:
            if name in parts and name not in names:
                names.append(name)
    return names


def apply_replacements(parts: dict[str, bytes], rules: list[Any], summary: dict[str, Any]) -> None:
    if not rules:
        return
    if not isinstance(rules, list):
        raise ValueError("replacements must be an array")
    report: list[dict[str, Any]] = []
    for rule_index, rule in enumerate(rules):
        if not isinstance(rule, dict) or "find" not in rule:
            raise ValueError("Each replacement needs find and optional replace")
        pattern = str(rule["find"])
        replacement = str(rule.get("replace", ""))
        regex = bool(rule.get("regex", False))
        replace_all = rule.get("occurrence", "all") != "first"
        total = 0
        for part_name in selected_part_names(parts, rule.get("parts")):
            document = parse_xml(parts[part_name])
            part_count = 0
            for paragraph in document.getElementsByTagName("w:p"):
                count = replace_text_in_paragraph(paragraph, pattern, replacement, regex, replace_all)
                total += count
                part_count += count
                if count and not replace_all:
                    break
            if part_count:
                parts[part_name] = xml_bytes(document)
            if total and not replace_all:
                break
        if rule.get("required") and total == 0:
            raise ValueError(f"replacements[{rule_index}] required text was not found: {pattern}")
        report.append({"rule": rule_index, "replacements": total})
    summary["replacements"] = report


def create_text_run(document: minidom.Document, text: str, run_config: dict[str, Any] | None = None) -> Node:
    run = document.createElement("w:r")
    if run_config:
        apply_fonts(ensure_direct(run, "w:rPr", RPR_ORDER), run_config)
    text_node = document.createElement("w:t")
    set_text_value(text_node, text)
    run.appendChild(text_node)
    return run


def validate_field_instruction(instruction: str) -> None:
    match = re.match(r"\s*([A-Za-z]+)", instruction)
    kind = match.group(1).upper() if match else ""
    if kind not in ALLOWED_FIELD_KINDS:
        raise ValueError(f"Unsupported or unsafe Word field instruction: {instruction}")


def create_field_runs(
    document: minidom.Document,
    instruction: str,
    result: str = "0",
    run_config: dict[str, Any] | None = None,
) -> list[Node]:
    validate_field_instruction(instruction)
    nodes: list[Node] = []
    begin = document.createElement("w:r")
    begin_char = document.createElement("w:fldChar")
    begin_char.setAttribute("w:fldCharType", "begin")
    begin_char.setAttribute("w:dirty", "true")
    begin.appendChild(begin_char)
    nodes.append(begin)
    instruction_run = document.createElement("w:r")
    instruction_node = document.createElement("w:instrText")
    instruction_node.setAttribute("xml:space", "preserve")
    instruction_node.appendChild(document.createTextNode(f" {instruction.strip()} "))
    instruction_run.appendChild(instruction_node)
    nodes.append(instruction_run)
    separate = document.createElement("w:r")
    separate_char = document.createElement("w:fldChar")
    separate_char.setAttribute("w:fldCharType", "separate")
    separate.appendChild(separate_char)
    nodes.append(separate)
    nodes.append(create_text_run(document, result, run_config))
    end = document.createElement("w:r")
    end_char = document.createElement("w:fldChar")
    end_char.setAttribute("w:fldCharType", "end")
    end.appendChild(end_char)
    nodes.append(end)
    return nodes


def apply_bottom_border(p_pr: Node, config: dict[str, Any]) -> None:
    if not config:
        return
    borders = ensure_direct(p_pr, "w:pBdr", PPR_ORDER)
    bottom = ensure_direct(borders, "w:bottom", ["w:top", "w:left", "w:bottom", "w:right", "w:between", "w:bar"])
    style = str(config.get("style", "single"))
    bottom.setAttribute("w:val", style)
    bottom.setAttribute("w:sz", str(config.get("size", 6)))
    bottom.setAttribute("w:space", str(config.get("space", 1)))
    bottom.setAttribute("w:color", str(config.get("color", "000000")).lstrip("#").upper())


def build_marginal_paragraph(document: minidom.Document, config: dict[str, Any]) -> Node:
    paragraph = document.createElement("w:p")
    p_pr = ensure_direct(paragraph, "w:pPr", ["w:pPr", "w:r"])
    paragraph_config = config.get("paragraph", {})
    if "alignment" in config:
        paragraph_config = {**paragraph_config, "alignment": config["alignment"]}
    run_config = config.get("run", {})
    if paragraph_config:
        apply_paragraph_format(p_pr, paragraph_config, run_config.get("sizePt"))
    apply_bottom_border(p_pr, config.get("bottomBorder", {}))
    segments = config.get("segments")
    if segments is None:
        segments = [{"text": str(config.get("text", ""))}]
    if not isinstance(segments, list):
        raise ValueError("header/footer paragraph segments must be an array")
    for segment in segments:
        if not isinstance(segment, dict):
            raise ValueError("header/footer segment must be an object")
        segment_run = {**run_config, **segment.get("run", {})}
        if "text" in segment:
            paragraph.appendChild(create_text_run(document, str(segment["text"]), segment_run))
        elif "field" in segment:
            instruction = str(segment["field"])
            for node in create_field_runs(document, instruction, str(segment.get("result", "0")), segment_run):
                paragraph.appendChild(node)
        else:
            raise ValueError("header/footer segment needs text or field")
    return paragraph


def next_marginal_part(parts: dict[str, bytes], kind: str) -> str:
    numbers = []
    pattern = re.compile(rf"word/{kind}(\d+)\.xml")
    for name in parts:
        match = pattern.fullmatch(name)
        if match:
            numbers.append(int(match.group(1)))
    return f"word/{kind}{max(numbers, default=0) + 1}.xml"


def set_even_odd_headers(parts: dict[str, bytes], enabled: bool) -> None:
    document = ensure_settings(parts)
    root = document.documentElement
    node = direct_child(root, "w:evenAndOddHeaders")
    if node is None:
        node = document.createElement("w:evenAndOddHeaders")
        before = direct_child(root, "w:updateFields")
        if before is None:
            before = next((child for child in element_children(root) if child.nodeName in SETTINGS_AFTER_UPDATE), None)
        if before is not None:
            root.insertBefore(node, before)
        else:
            root.appendChild(node)
    node.setAttribute("w:val", "true" if enabled else "false")
    parts["word/settings.xml"] = xml_bytes(document)


def remove_marginal_reference(section: Node, kind: str, reference_type: str) -> None:
    tag = f"w:{kind}Reference"
    for node in list(element_children(section)):
        if node.nodeName == tag and node.getAttribute("w:type") == reference_type:
            section.removeChild(node)


def apply_headers_footers(
    parts: dict[str, bytes], document: minidom.Document, rules: list[Any], summary: dict[str, Any]
) -> None:
    if not rules:
        return
    if not isinstance(rules, list):
        raise ValueError("headersFooters must be an array")
    if not document.documentElement.hasAttribute("xmlns:r"):
        document.documentElement.setAttribute("xmlns:r", OFFICE_REL)
    sections = document_sections(document)
    report: list[dict[str, Any]] = []
    for index, rule in enumerate(rules):
        if not isinstance(rule, dict):
            raise ValueError("Each headersFooters item must be an object")
        section_index = rule.get("section")
        kind = rule.get("kind")
        reference_type = rule.get("type", "default")
        action = rule.get("action", "set")
        if not isinstance(section_index, int) or not 0 <= section_index < len(sections):
            raise ValueError(f"headersFooters[{index}] has invalid section")
        if kind not in {"header", "footer"}:
            raise ValueError("headersFooters.kind must be header or footer")
        if reference_type not in HEADER_FOOTER_TYPES:
            raise ValueError("headersFooters.type must be default, first, or even")
        if action not in {"set", "clear", "inherit"}:
            raise ValueError("headersFooters.action must be set, clear, or inherit")
        section = sections[section_index]
        if action == "inherit":
            report.append({"rule": index, "action": action})
            continue
        remove_marginal_reference(section, kind, reference_type)
        part_name = next_marginal_part(parts, kind)
        root_tag = "w:hdr" if kind == "header" else "w:ftr"
        marginal = minidom.parseString(
            f'<?xml version="1.0" encoding="UTF-8"?><{root_tag} xmlns:w="{W_NS}"/>'
        )
        root = marginal.documentElement
        paragraphs = rule.get("paragraphs", []) if action == "set" else []
        if action == "set" and not paragraphs:
            paragraphs = [{"text": str(rule.get("text", "")), "alignment": rule.get("alignment", "center"),
                           "run": rule.get("run", {}), "segments": rule.get("segments"),
                           "bottomBorder": rule.get("bottomBorder", {})}]
        if not isinstance(paragraphs, list):
            raise ValueError("headersFooters.paragraphs must be an array")
        for paragraph_config in paragraphs or [{}]:
            root.appendChild(build_marginal_paragraph(marginal, paragraph_config))
        parts[part_name] = xml_bytes(marginal)
        relationship_id = ensure_relationship(
            parts, HEADER_REL if kind == "header" else FOOTER_REL, part_name.removeprefix("word/")
        )
        ensure_content_type(parts, part_name, HEADER_CT if kind == "header" else FOOTER_CT)
        reference = document.createElement(f"w:{kind}Reference")
        reference.setAttribute("w:type", reference_type)
        reference.setAttribute("r:id", relationship_id)
        insert_ordered(section, reference, SECTPR_ORDER)
        if reference_type == "first":
            set_toggle(section, "w:titlePg", True, SECTPR_ORDER)
        if reference_type == "even":
            set_even_odd_headers(parts, True)
        report.append({"rule": index, "action": action, "part": part_name, "relationshipId": relationship_id})
    summary["headersFooters"] = report


def replace_placeholder_with_field(
    paragraph: Node, placeholder: str, instruction: str, result: str, run_config: dict[str, Any]
) -> bool:
    full_text = paragraph_text(paragraph)
    start = full_text.find(placeholder)
    if start < 0:
        return False
    end = start + len(placeholder)
    nodes = paragraph_text_nodes(paragraph)
    offset = 0
    first_node = None
    last_node = None
    first_left = last_left = 0
    for node in nodes:
        value = text_value(node)
        right = offset + len(value)
        if first_node is None and offset <= start < right:
            first_node, first_left = node, offset
        if offset < end <= right:
            last_node, last_left = node, offset
            break
        offset = right
    if first_node is None or last_node is None:
        return False
    first_run = ancestor(first_node, "w:r")
    if first_run is None or ancestor(first_run, "w:p") is not paragraph:
        return False
    parent = first_run.parentNode
    if parent is None:
        return False
    first_value = text_value(first_node)
    last_value = text_value(last_node)
    prefix = first_value[: start - first_left]
    suffix = last_value[end - last_left :]
    set_text_value(first_node, prefix)
    if first_node is not last_node:
        reached = False
        for node in nodes:
            if node is first_node:
                reached = True
                continue
            if not reached:
                continue
            if node is last_node:
                set_text_value(node, suffix)
                break
            set_text_value(node, "")
    insertion_point = first_run.nextSibling
    for field_run in create_field_runs(paragraph.ownerDocument, instruction, result, run_config):
        parent.insertBefore(field_run, insertion_point)
    if first_node is last_node and suffix:
        parent.insertBefore(create_text_run(paragraph.ownerDocument, suffix, run_config), insertion_point)
    return True


def apply_field_rules(parts: dict[str, bytes], rules: list[Any], summary: dict[str, Any]) -> None:
    if not rules:
        return
    if not isinstance(rules, list):
        raise ValueError("fieldRules must be an array")
    report: list[dict[str, Any]] = []
    for index, rule in enumerate(rules):
        if not isinstance(rule, dict) or "placeholder" not in rule or "instruction" not in rule:
            raise ValueError("Each fieldRules item needs placeholder and instruction")
        placeholder = str(rule["placeholder"])
        instruction = str(rule["instruction"])
        validate_field_instruction(instruction)
        total = 0
        replace_all = rule.get("occurrence", "all") != "first"
        for part_name in selected_part_names(parts, rule.get("parts", ["document"])):
            document = parse_xml(parts[part_name])
            changed = 0
            for paragraph in document.getElementsByTagName("w:p"):
                while replace_placeholder_with_field(
                    paragraph, placeholder, instruction, str(rule.get("result", "0")), rule.get("run", {})
                ):
                    changed += 1
                    total += 1
                    if not replace_all:
                        break
                if changed and not replace_all:
                    break
            if changed:
                parts[part_name] = xml_bytes(document)
            if total and not replace_all:
                break
        if rule.get("required") and total == 0:
            raise ValueError(f"fieldRules[{index}] required placeholder was not found: {placeholder}")
        report.append({"rule": index, "fieldsInserted": total, "instruction": instruction})
    summary["fieldRules"] = report


def next_bookmark_id(document: minidom.Document) -> int:
    ids = []
    for node in document.getElementsByTagName("w:bookmarkStart"):
        value = node.getAttribute("w:id")
        if value.isdigit():
            ids.append(int(value))
    return max(ids, default=-1) + 1


def add_bookmark(paragraph: Node, name: str, bookmark_id: int) -> None:
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,39}", name):
        raise ValueError(f"Invalid Word bookmark name: {name}")
    document = paragraph.ownerDocument
    start = document.createElement("w:bookmarkStart")
    start.setAttribute("w:id", str(bookmark_id))
    start.setAttribute("w:name", name)
    end = document.createElement("w:bookmarkEnd")
    end.setAttribute("w:id", str(bookmark_id))
    p_pr = direct_child(paragraph, "w:pPr")
    insertion = p_pr.nextSibling if p_pr is not None else paragraph.firstChild
    paragraph.insertBefore(start, insertion)
    paragraph.appendChild(end)


def apply_bookmark_rules(document: minidom.Document, rules: list[Any], summary: dict[str, Any]) -> None:
    if not rules:
        return
    if not isinstance(rules, list):
        raise ValueError("bookmarkRules must be an array")
    infos = paragraph_infos(document)
    bookmark_id = next_bookmark_id(document)
    existing = {node.getAttribute("w:name") for node in document.getElementsByTagName("w:bookmarkStart")}
    report: list[dict[str, Any]] = []
    for rule_index, rule in enumerate(rules):
        if not isinstance(rule, dict) or "name" not in rule:
            raise ValueError("Each bookmarkRules item needs name")
        matched = [info for info in infos if paragraph_matches(info, rule.get("match", {}))]
        if rule.get("required") and not matched:
            raise ValueError(f"bookmarkRules[{rule_index}] required match was not found")
        max_matches = rule.get("maxMatches", len(matched))
        if not isinstance(max_matches, int) or max_matches < 1:
            raise ValueError("bookmarkRules.maxMatches must be a positive integer")
        created = []
        for occurrence, info in enumerate(matched[:max_matches], start=1):
            name = str(rule["name"]).format(index=occurrence, paragraph=info["index"])
            if name in existing:
                continue
            add_bookmark(info["paragraph"], name, bookmark_id)
            existing.add(name)
            created.append(name)
            bookmark_id += 1
        report.append({"rule": rule_index, "created": created})
    summary["bookmarkRules"] = report


def clear_paragraph_content(paragraph: Node) -> None:
    p_pr = direct_child(paragraph, "w:pPr")
    for child in list(paragraph.childNodes):
        if child is not p_pr:
            paragraph.removeChild(child)


def caption_chapter_instruction(profile: dict[str, Any], rule: dict[str, Any]) -> str | None:
    if rule.get("chapterFieldInstruction"):
        return str(rule["chapterFieldInstruction"])
    numbering = profile.get("captionNumbering") or {}
    if not isinstance(numbering, dict):
        return None
    if numbering.get("chapterFieldInstruction"):
        return str(numbering["chapterFieldInstruction"])
    strategy = numbering.get("strategy")
    style = numbering.get("chapterStyle") or "Heading 1"
    if strategy == "styleref-as-displayed":
        return f'STYLEREF "{style}" \\n'
    if strategy == "styleref-arabic":
        return f'STYLEREF "{style}" \\n \\* ARABIC'
    if strategy == "seq-chapter":
        return "SEQ chapter \\* ARABIC"
    return None


def apply_caption_rules(
    document: minidom.Document,
    rules: list[Any],
    default_run: dict[str, Any],
    summary: dict[str, Any],
    profile: dict[str, Any] | None = None,
) -> None:
    if not rules:
        return
    if not isinstance(rules, list):
        raise ValueError("captionRules must be an array")
    infos = paragraph_infos(document)
    report: list[dict[str, Any]] = []
    bookmark_id = next_bookmark_id(document)
    profile = profile or {}
    numbering = profile.get("captionNumbering") or {}
    default_separator = numbering.get("separator") if isinstance(numbering, dict) else None
    for rule_index, rule in enumerate(rules):
        if not isinstance(rule, dict) or "label" not in rule:
            raise ValueError("Each captionRules item needs label")
        matched = [info for info in infos if paragraph_matches(info, rule.get("match", {}))]
        if rule.get("required") and not matched:
            raise ValueError(f"captionRules[{rule_index}] required match was not found")
        count = 0
        for occurrence, info in enumerate(matched, start=1):
            paragraph = info["paragraph"]
            original = paragraph_text(paragraph)
            strip_pattern = rule.get("stripPattern")
            remainder = re.sub(str(strip_pattern), "", original, count=1) if strip_pattern else original
            clear_paragraph_content(paragraph)
            p_pr = direct_child(paragraph, "w:pPr") or ensure_direct(paragraph, "w:pPr", ["w:pPr", "w:r"])
            if "style" in rule:
                set_val(ensure_direct(p_pr, "w:pStyle", PPR_ORDER), str(rule["style"]))
            run_config = {**default_run, **rule.get("run", {})}
            label = str(rule["label"])
            paragraph.appendChild(create_text_run(document, label, run_config))
            chapter_instruction = caption_chapter_instruction(profile, rule)
            if chapter_instruction:
                for node in create_field_runs(document, str(chapter_instruction), str(rule.get("chapterResult", "1")), run_config):
                    paragraph.appendChild(node)
                paragraph.appendChild(create_text_run(document, str(rule.get("separator", default_separator or "-")), run_config))
            sequence = str(rule.get("sequence", label))
            sequence_instruction = str(rule.get("sequenceInstruction", f"SEQ {sequence} \\* ARABIC"))
            for node in create_field_runs(document, sequence_instruction, str(rule.get("sequenceResult", occurrence)), run_config):
                paragraph.appendChild(node)
            paragraph.appendChild(create_text_run(document, str(rule.get("afterNumber", " ")) + remainder.lstrip(), run_config))
            bookmark_prefix = rule.get("bookmarkPrefix")
            if bookmark_prefix:
                name = f"{bookmark_prefix}{occurrence}"
                add_bookmark(paragraph, name, bookmark_id)
                bookmark_id += 1
            count += 1
        report.append({"rule": rule_index, "captions": count})
    summary["captionRules"] = report


def check_requirements_metadata(profile: dict[str, Any], allow_unresolved: bool) -> None:
    requirements = profile.get("requirements", {})
    if not requirements:
        return
    if not isinstance(requirements, dict):
        raise ValueError("requirements must be an object")
    unresolved = requirements.get("unresolved", [])
    conflicts = requirements.get("conflicts", [])
    if not isinstance(unresolved, list) or not isinstance(conflicts, list):
        raise ValueError("requirements.unresolved and requirements.conflicts must be arrays")
    if (unresolved or conflicts) and not allow_unresolved:
        raise ValueError(
            "Profile still has unresolved requirements or conflicts; resolve them or pass --allow-unresolved"
        )


def validate_package_parts(parts: dict[str, bytes]) -> None:
    for required in ("[Content_Types].xml", "word/document.xml"):
        if required not in parts:
            raise ValueError(f"Missing required DOCX part: {required}")
    for name, data in parts.items():
        if name.endswith(".xml") or name.endswith(".rels"):
            parse_xml(data)


def write_package(
    source_infos: list[zipfile.ZipInfo], parts: dict[str, bytes], output: Path, force: bool
) -> None:
    if output.exists() and not force:
        raise FileExistsError(f"Output exists: {output}; pass --force to replace it")
    output.parent.mkdir(parents=True, exist_ok=True)
    existing_names = {info.filename for info in source_infos}
    if len(existing_names) != len(source_infos):
        raise ValueError("Input DOCX contains duplicate ZIP entry names")
    handle = tempfile.NamedTemporaryFile(prefix=output.stem + ".", suffix=".tmp", dir=output.parent, delete=False)
    temp_path = Path(handle.name)
    handle.close()
    try:
        with zipfile.ZipFile(temp_path, "w") as archive:
            for info in source_infos:
                archive.writestr(info, parts[info.filename])
            for name in sorted(parts.keys() - existing_names):
                archive.writestr(name, parts[name], compress_type=zipfile.ZIP_DEFLATED)
        with zipfile.ZipFile(temp_path, "r") as archive:
            corrupt = archive.testzip()
            if corrupt:
                raise ValueError(f"Generated package has corrupt member: {corrupt}")
        os.replace(temp_path, output)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def apply_profile(
    input_path: Path,
    output_path: Path,
    profile_path: Path | None,
    force: bool,
    allow_unresolved: bool = False,
    *,
    profile_data: dict[str, Any] | None = None,
    skip_schema: bool = False,
) -> dict[str, Any]:
    if input_path.resolve() == output_path.resolve():
        raise ValueError("Refusing in-place write: input and output paths are identical")
    if profile_data is not None:
        profile = profile_data
    elif profile_path is not None:
        profile = json.loads(profile_path.read_text(encoding="utf-8"))
    else:
        raise ValueError("A profile path or profile_data object is required")
    if not isinstance(profile, dict):
        raise ValueError("Profile root must be an object")
    if not skip_schema:
        ensure_profile(profile)
    check_requirements_metadata(profile, allow_unresolved)

    with zipfile.ZipFile(input_path, "r") as archive:
        source_infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in source_infos}
        corrupt = archive.testzip()
        if corrupt:
            raise ValueError(f"Input DOCX has corrupt member: {corrupt}")
    validate_package_parts(parts)

    summary: dict[str, Any] = {"profile": profile.get("name"), "input": str(input_path.resolve())}
    document = parse_xml(parts["word/document.xml"])
    styles_document, styles_root = ensure_styles_root(parts)
    default_run = profile.get("defaultRun", {})
    if not isinstance(default_run, dict):
        raise ValueError("defaultRun must be an object")

    apply_page(document, profile.get("page", {}), summary)
    apply_section_configs(document, profile.get("sections", []), summary)
    apply_default_run(styles_root, default_run)
    summary["stylesConfigured"] = apply_styles(styles_root, profile.get("styles", {}), default_run)
    apply_heading_numbering(parts, styles_root, profile.get("headingNumbering", {}), summary)
    apply_page_numbers(document, profile.get("pageNumbers", []), summary)
    style_names = style_id_to_name(parts)
    apply_paragraph_rules(document, profile.get("paragraphRules", []), default_run, summary, style_names=style_names)
    apply_caption_rules(document, profile.get("captionRules", []), default_run, summary, profile)
    apply_bookmark_rules(document, profile.get("bookmarkRules", []), summary)
    apply_headers_footers(parts, document, profile.get("headersFooters", []), summary)
    parts["word/document.xml"] = xml_bytes(document)
    apply_replacements(parts, profile.get("replacements", []), summary)
    apply_field_rules(parts, profile.get("fieldRules", []), summary)
    fields_config = profile.get("fields", {})
    if not isinstance(fields_config, dict):
        raise ValueError("fields must be an object")
    apply_update_fields(parts, fields_config, summary)
    apply_theme(parts, default_run, summary)

    parts["word/styles.xml"] = xml_bytes(styles_document)
    ensure_relationship(parts, STYLES_REL, "styles.xml")
    ensure_content_type(parts, "word/styles.xml", STYLES_CT)
    validate_package_parts(parts)
    write_package(source_infos, parts, output_path, force)
    summary["ok"] = True
    summary["output"] = str(output_path.resolve())
    if profile.get("captionNumbering"):
        summary["captionNumbering"] = profile.get("captionNumbering")
    return summary


def parse_id_list(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def apply_with_optional_plan(
    input_path: Path,
    output_path: Path,
    profile_path: Path,
    force: bool,
    allow_unresolved: bool,
    plan_path: Path | None = None,
    only_ids: list[str] | None = None,
    skip_ids: list[str] | None = None,
    only_auto: bool = False,
    include_page_breaks: bool = False,
    confirmed: bool = False,
) -> dict[str, Any]:
    from plan_apply import filter_profile_for_steps, select_steps, visual_issues_from_steps
    from visual_repair import apply_visual_repairs

    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    selected_steps: list[dict[str, Any]] = []
    if plan_path is not None:
        if not confirmed:
            raise ValueError("Applying a repair plan requires --yes after the user confirms the selected steps")
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
        if not isinstance(plan, dict) or not isinstance(plan.get("repairPlan"), list):
            raise ValueError("Repair plan JSON must contain a repairPlan array")
        selected_steps = select_steps(
            plan,
            only_ids=only_ids,
            skip_ids=skip_ids,
            only_auto=only_auto,
            include_page_breaks=include_page_breaks,
        )
        if not selected_steps:
            raise ValueError("No repair-plan steps selected; pass --only-auto, --only-ids, or confirm the full auto set")
        profile = filter_profile_for_steps(profile, selected_steps)
        # A confirmed plan is a scoped mutation; leftover unresolved clauses
        # that were not selected must not block the chosen steps.
        allow_unresolved = True

    result = apply_profile(
        input_path,
        output_path,
        profile_path,
        force,
        allow_unresolved,
        profile_data=profile,
    )
    visual_issues = visual_issues_from_steps(selected_steps)
    if visual_issues:
        visual_output = output_path.with_name(output_path.stem + ".visual.docx")
        visual_result = apply_visual_repairs(output_path, visual_output, visual_issues, force=True)
        result["visualRepair"] = visual_result
        result["output"] = visual_result["output"]
    if selected_steps:
        result["appliedStepIds"] = [step.get("id") for step in selected_steps]
        result["appliedStepCount"] = len(selected_steps)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply a DOCX formatting profile to a new output file.")
    parser.add_argument("--input", required=True, type=Path, help="Source .docx")
    parser.add_argument("--output", required=True, type=Path, help="New standardized .docx")
    parser.add_argument("--profile", required=True, type=Path, help="UTF-8 JSON profile")
    parser.add_argument("--force", action="store_true", help="Replace an existing output file")
    parser.add_argument(
        "--allow-unresolved",
        action="store_true",
        help="Apply even when profile requirements list unresolved items or conflicts",
    )
    parser.add_argument("--plan", type=Path, help="Confirmed repair-plan JSON; apply only selected steps")
    parser.add_argument("--only-ids", help="Comma-separated repair step ids to apply")
    parser.add_argument("--skip-ids", help="Comma-separated repair step ids to skip")
    parser.add_argument("--only-auto", action="store_true", help="Apply only auto-applyable plan steps")
    parser.add_argument("--include-page-breaks", action="store_true", help="Allow set_page_break_before steps")
    parser.add_argument("--yes", action="store_true", help="Confirm applying a repair plan")
    args = parser.parse_args()
    try:
        for path, label in ((args.input, "input"), (args.profile, "profile")):
            if not path.is_file():
                raise FileNotFoundError(f"Missing {label}: {path}")
        result = apply_with_optional_plan(
            args.input,
            args.output,
            args.profile,
            args.force,
            args.allow_unresolved,
            plan_path=args.plan,
            only_ids=parse_id_list(args.only_ids),
            skip_ids=parse_id_list(args.skip_ids),
            only_auto=args.only_auto,
            include_page_breaks=args.include_page_breaks,
            confirmed=args.yes,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
