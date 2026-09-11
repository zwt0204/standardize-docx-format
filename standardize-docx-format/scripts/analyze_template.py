#!/usr/bin/env python3
"""Analyze an official DOCX template and emit a draft profile."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

from audit_docx import attr, audit, direct_child, parse_xml
from compile_requirements import DEFAULT_THESIS_SECTIONS
from document_model import build_model
from diff_profile import PAGE_SIZES_IN
from profile_schema import ensure_profile

NUMBER_FORMAT_REVERSE = {
    "decimal": "decimal",
    "lowerRoman": "lower-roman",
    "upperRoman": "upper-roman",
    "lowerLetter": "lower-alpha",
    "upperLetter": "upper-alpha",
    "chineseCounting": "chinese-counting",
    "chineseCountingThousand": "chinese-counting-thousand",
    "chineseLegalSimplified": "chinese-legal-simplified",
    "ideographTraditional": "ideograph-traditional",
    "ideographDigital": "ideograph-digital",
    "decimalEnclosedCircle": "decimal-enclosed-circle",
    "decimalEnclosedFullstop": "decimal-enclosed-fullstop",
    "decimalFullWidth": "decimal-full-width",
}

HEADING_STYLE_IDS = {"Heading1", "Heading2", "Heading3", "Heading4", "Heading5"}
PLACEHOLDER_RE = re.compile(r"\{\{[^{}]+\}\}|＜[^＞]{2,40}＞|【[^】]{2,40}】")
TOC_SAFE = {"styleNameNotRegex": "(?i)^toc", "textNotRegex": r".+\d$", "inTextBox": False}


def closest_page_size(width: float | None, height: float | None) -> str | None:
    if width is None or height is None:
        return None
    best = None
    best_delta = 1e9
    for name, (w, h) in PAGE_SIZES_IN.items():
        delta = min(abs(width - w) + abs(height - h), abs(width - h) + abs(height - w))
        if delta < best_delta:
            best = name
            best_delta = delta
    return best if best_delta <= 0.15 else None


def compact(mapping: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in mapping.items() if value not in (None, {}, [])}


def extract_heading_numbering(parts: dict[str, bytes], _styles: dict[str, Any] | None = None) -> dict[str, Any] | None:
    if "word/numbering.xml" not in parts:
        return None
    numbering = parse_xml(parts["word/numbering.xml"])
    style_to_num: dict[str, tuple[str, str]] = {}
    if "word/styles.xml" in parts:
        styles_dom = parse_xml(parts["word/styles.xml"])
        for style in styles_dom.getElementsByTagName("w:style"):
            style_id = attr(style, "w:styleId") or ""
            name = (attr(direct_child(style, "w:name"), "w:val") or "").lower()
            num_pr = direct_child(direct_child(style, "w:pPr"), "w:numPr")
            if not style_id or num_pr is None:
                continue
            if style_id not in HEADING_STYLE_IDS and not name.startswith("heading"):
                continue
            num_id = attr(direct_child(num_pr, "w:numId"), "w:val") or ""
            ilvl = attr(direct_child(num_pr, "w:ilvl"), "w:val") or "0"
            style_to_num[style_id] = (num_id, ilvl)

    abstract_by_num: dict[str, str] = {}
    for num in numbering.getElementsByTagName("w:num"):
        num_id = attr(num, "w:numId")
        abstract_id = attr(direct_child(num, "w:abstractNumId"), "w:val")
        if num_id and abstract_id:
            abstract_by_num[num_id] = abstract_id
    abstracts: dict[str, dict[str, dict[str, str]]] = {}
    for abstract in numbering.getElementsByTagName("w:abstractNum"):
        abstract_id = attr(abstract, "w:abstractNumId")
        levels: dict[str, dict[str, str]] = {}
        for lvl in abstract.getElementsByTagName("w:lvl"):
            ilvl = attr(lvl, "w:ilvl") or "0"
            levels[ilvl] = {
                "pattern": attr(direct_child(lvl, "w:lvlText"), "w:val") or "",
                "format": attr(direct_child(lvl, "w:numFmt"), "w:val") or "decimal",
                "start": attr(direct_child(lvl, "w:start"), "w:val") or "1",
                "pStyle": attr(direct_child(lvl, "w:pStyle"), "w:val") or "",
                "suffix": attr(direct_child(lvl, "w:suff"), "w:val") or "tab",
            }
        abstracts[abstract_id or ""] = levels

    chosen_levels: list[tuple[int, str, dict[str, str]]] = []
    for style_id, (num_id, ilvl) in style_to_num.items():
        abstract_id = abstract_by_num.get(str(num_id))
        level = (abstracts.get(abstract_id or {}) or {}).get(str(ilvl))
        if level:
            chosen_levels.append((int(ilvl), style_id, level))
    if not chosen_levels:
        for _abstract_id, levels in abstracts.items():
            heading_levels = [
                (int(ilvl), data.get("pStyle") or f"Heading{int(ilvl) + 1}", data)
                for ilvl, data in levels.items()
                if data.get("pStyle") in HEADING_STYLE_IDS or str(data.get("pStyle") or "").startswith("Heading")
            ]
            if heading_levels:
                chosen_levels = heading_levels
                break
    if not chosen_levels:
        return None
    chosen_levels.sort(key=lambda item: item[0])
    styles_out = [item[1] for item in chosen_levels]
    patterns = [item[2]["pattern"] for item in chosen_levels]
    formats = [NUMBER_FORMAT_REVERSE.get(item[2]["format"], "decimal") for item in chosen_levels]
    starts = [int(item[2]["start"]) if str(item[2]["start"]).isdigit() else 1 for item in chosen_levels]
    suffix_raw = chosen_levels[0][2].get("suffix") or "space"
    suffix = {"nothing": "nothing", "tab": "tab"}.get(suffix_raw, "space")
    if not all(f"%{index}" in pattern for index, pattern in enumerate(patterns, start=1)):
        return None
    return {
        "enabled": True,
        "styles": styles_out,
        "patterns": patterns,
        "formats": formats,
        "starts": starts,
        "suffix": suffix,
    }


def extract_field_rules(instructions: list[str]) -> tuple[list[dict[str, Any]], list[str]]:
    field_rules: list[dict[str, Any]] = []
    unresolved: list[str] = []
    seen: set[str] = set()
    for instruction in instructions:
        kind_match = re.match(r"\s*([A-Za-z]+)", instruction)
        kind = (kind_match.group(1).upper() if kind_match else "")
        if kind == "TOC" and "TOC" not in seen:
            seen.add("TOC")
            field_rules.append(
                {
                    "placeholder": "{{TOC}}",
                    "instruction": instruction,
                    "result": "目录",
                    "required": False,
                }
            )
        elif kind in {"CITATION", "BIBLIOGRAPHY"} and kind not in seen:
            seen.add(kind)
            unresolved.append(f"模板包含 {kind} 字段；学校引用体系仍需确认，不自动改写文献")
        elif kind == "SEQ" and "SEQ" not in seen:
            seen.add("SEQ")
            unresolved.append(f"模板包含 SEQ 字段（{instruction[:80]}），题注编号策略需与 captionNumbering 对齐")
        elif kind == "STYLEREF" and "STYLEREF" not in seen:
            seen.add("STYLEREF")
            unresolved.append("模板使用 STYLEREF 取章号；若正文是中文序号而图表要阿拉伯数字，需确认 captionNumbering.strategy")
    return field_rules, unresolved


def extract_placeholders(parts: dict[str, bytes]) -> list[dict[str, Any]]:
    found: dict[str, set[str]] = {}
    for name, data in parts.items():
        if not name.endswith(".xml"):
            continue
        try:
            text = data.decode("utf-8", errors="ignore")
        except Exception:
            continue
        for match in PLACEHOLDER_RE.findall(text):
            bucket = "document"
            if re.fullmatch(r"word/header\d+\.xml", name):
                bucket = "headers"
            elif re.fullmatch(r"word/footer\d+\.xml", name):
                bucket = "footers"
            found.setdefault(match, set()).add(bucket)
    replacements = []
    for token, parts_used in sorted(found.items()):
        replacements.append(
            {
                "find": token,
                "replace": "",
                "parts": sorted(parts_used),
                "occurrence": "all",
                "required": False,
                "note": "Draft placeholder from the official template; fill replace after confirming the cover field.",
            }
        )
    return replacements


def infer_paragraph_rules(model: dict[str, Any]) -> list[dict[str, Any]]:
    rules: list[dict[str, Any]] = []
    chapter_like = 0
    section_like = 0
    chinese_chapter = 0
    for block in model.get("blocks") or []:
        if block.get("type") == "toc_entry":
            continue
        text = str(block.get("preview") or block.get("text") or "").strip()
        if re.match(r"^[一二三四五六七八九十]+、", text):
            chinese_chapter += 1
        if re.match(r"^\d+\s+\S+", text) and block.get("type") == "chapter":
            chapter_like += 1
        if re.match(r"^\d+\.\d+\s+\S+", text) and block.get("type") == "section":
            section_like += 1
    if chinese_chapter:
        rules.append(
            {
                "match": {
                    "textRegex": r"^[一二三四五六七八九十]+、",
                    "styleNameNotRegex": "(?i)^toc",
                    "inTextBox": False,
                },
                "style": "Heading1",
                "textRegexReplace": {"pattern": r"^[一二三四五六七八九十]+、\s*", "replacement": ""},
                "maxMatches": 50,
            }
        )
    elif chapter_like:
        rules.append(
            {
                "match": {
                    "textRegex": r"^\d+\s+\S+",
                    **TOC_SAFE,
                },
                "style": "Heading1",
                "textRegexReplace": {"pattern": r"^\d+\s+", "replacement": ""},
                "maxMatches": 50,
            }
        )
    if section_like:
        rules.append(
            {
                "match": {
                    "textRegex": r"^\d+\.\d+\s+\S+",
                    **TOC_SAFE,
                },
                "style": "Heading2",
                "textRegexReplace": {"pattern": r"^\d+\.\d+\s+", "replacement": ""},
                "maxMatches": 100,
            }
        )
    return rules


def detect_caption_strategy(heading_numbering: dict[str, Any] | None, unresolved: list[str]) -> dict[str, Any]:
    formats = (heading_numbering or {}).get("formats") or []
    heading_is_chinese = bool(formats) and str(formats[0]).startswith("chinese")
    caption = {
        "chapterStyle": "Heading 1",
        "separator": "-",
        "strategy": "styleref-arabic" if heading_is_chinese else "styleref-as-displayed",
    }
    if heading_is_chinese:
        unresolved.append(
            "一级标题使用中文序号而图表通常要求阿拉伯章号；草稿采用 styleref-arabic，应用前请确认学校题注策略"
        )
        caption["headingDisplay"] = "chinese-counting"
        caption["chapterDisplay"] = "decimal"
    return caption


def analyze_template(path: Path) -> dict[str, Any]:
    report = audit(path)
    model = build_model(path, {"document": {"type": "thesis"}})
    with zipfile.ZipFile(path, "r") as archive:
        parts = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    sections = report.get("sections") or []
    first = (sections[0].get("page") if sections else {}) or {}
    size_name = closest_page_size(first.get("widthIn"), first.get("heightIn"))
    page: dict[str, Any] = compact(
        {
            "size": size_name,
            "orientation": str(first.get("orientation") or "portrait").replace("/default", "") or "portrait",
            "marginsIn": compact(first.get("marginsIn") or {}),
            "widthIn": None if size_name else first.get("widthIn"),
            "heightIn": None if size_name else first.get("heightIn"),
        }
    )

    default_run = compact(report.get("defaultRun") or {})
    profile_run = compact(
        {
            "latinFont": default_run.get("latin"),
            "eastAsiaFont": default_run.get("eastAsia"),
            "complexScriptFont": default_run.get("complexScript"),
            "sizePt": default_run.get("sizePt"),
            "updateTheme": True,
        }
    )

    interesting = {"Normal", "Heading1", "Heading2", "Heading3", "Heading4", "Caption", "TOC1", "TOC2"}
    styles_out: dict[str, Any] = {}
    usage = (report.get("content") or {}).get("styleUsage") or {}
    for style_id, style in (report.get("styles") or {}).items():
        if style_id not in interesting and not str(style_id).startswith("Heading") and usage.get(style_id, 0) < 3:
            continue
        if style.get("type") not in {None, "paragraph"}:
            continue
        run = style.get("run") or {}
        paragraph = style.get("paragraph") or {}
        styles_out[style_id] = compact(
            {
                "name": style.get("name"),
                "basedOn": style.get("basedOn"),
                "next": style.get("next"),
                "run": compact(
                    {
                        "latinFont": run.get("latin") or (style.get("fonts") or {}).get("latin"),
                        "eastAsiaFont": run.get("eastAsia") or (style.get("fonts") or {}).get("eastAsia"),
                        "complexScriptFont": run.get("complexScript") or (style.get("fonts") or {}).get("complexScript"),
                        "sizePt": run.get("sizePt") or style.get("sizePt"),
                        "bold": run.get("bold") if run.get("bold") else None,
                    }
                ),
                "paragraph": compact(
                    {
                        "alignment": paragraph.get("alignment"),
                        "lineSpacing": paragraph.get("lineSpacing"),
                        "spaceBeforePt": paragraph.get("spaceBeforePt"),
                        "spaceAfterPt": paragraph.get("spaceAfterPt"),
                        "firstLineIn": paragraph.get("firstLineIn"),
                        "keepNext": paragraph.get("keepNext"),
                        "widowControl": paragraph.get("widowControl"),
                        "outlineLevel": paragraph.get("outlineLevel"),
                    }
                ),
            }
        )

    page_numbers = []
    for section in sections:
        number = section.get("pageNumber") or {}
        if number.get("format") or number.get("start") is not None:
            reverse = {
                "decimal": "decimal",
                "lowerRoman": "lower-roman",
                "upperRoman": "upper-roman",
                "lowerLetter": "lower-alpha",
                "upperLetter": "upper-alpha",
            }
            page_numbers.append(
                compact(
                    {
                        "section": section.get("index"),
                        "format": reverse.get(number.get("format"), number.get("format")),
                        "start": number.get("start"),
                    }
                )
            )

    headers_footers = []
    for section in sections:
        for kind in ("header", "footer"):
            for reference in section.get(f"{kind}References") or []:
                headers_footers.append(
                    {
                        "section": section.get("index"),
                        "kind": kind,
                        "type": reference.get("type") or "default",
                        "action": "inherit",
                        "note": "Preserve the official template part; do not rebuild unless the spec requires it.",
                    }
                )

    unresolved = []
    if not profile_run.get("eastAsiaFont"):
        unresolved.append("模板未给出明确的默认东亚字体")
    if not page_numbers:
        unresolved.append("模板未包含可读取的分节页码格式")

    heading_numbering = extract_heading_numbering(parts, report.get("styles") or {})
    field_rules, field_unresolved = extract_field_rules((report.get("fields") or {}).get("instructions") or [])
    unresolved.extend(field_unresolved)
    replacements = extract_placeholders(parts)
    if replacements:
        unresolved.append("将封面现有填写文本或占位符逐项替换为稳定的双花括号占位符后，才能建立必填字段映射")
    else:
        unresolved.append("模板未检测到 {{PLACEHOLDER}} / 【】 / ＜＞ 占位符；精确封面应先在官方模板副本中建立稳定占位符")
    paragraph_rules = infer_paragraph_rules(model)
    caption_numbering = detect_caption_strategy(heading_numbering, unresolved)
    kinds = (report.get("fields") or {}).get("kinds") or {}
    if int(kinds.get("TOC") or 0) == 0:
        unresolved.append("模板未检测到 TOC 字段；目录页码必须由 Word 刷新，不要把目录结果写成普通文本")

    validation: dict[str, Any] = {
        "sectionCount": {"min": max(len(sections), 1)},
        "requiredStylesUsed": [style_id for style_id in ("Normal", "Heading1") if style_id in styles_out],
        "noPlaceholders": [r"\{\{[^{}]+\}\}"],
        "requiredTexts": ["摘要", "Abstract", "目录", "参考文献"],
    }
    if int(kinds.get("TOC") or 0) > 0:
        validation["requiredFields"] = {"TOC": 1}
    if int(kinds.get("PAGE") or 0) > 0:
        validation.setdefault("requiredFields", {})["PAGE"] = 1

    profile = {
        "name": f"从模板分析: {path.name}",
        "profileVersion": "2.1",
        "requirements": {
            "sourceType": "template",
            "sources": [path.name],
            "unresolved": unresolved,
            "conflicts": [],
        },
        "document": {
            "type": "thesis",
            "sections": DEFAULT_THESIS_SECTIONS,
            "observed": model.get("summary"),
        },
        "page": page,
        "defaultRun": profile_run,
        "styles": styles_out,
        "pageNumbers": page_numbers,
        "headersFooters": headers_footers,
        "fields": {"updateOnOpen": True},
        "validation": validation,
        "captionNumbering": caption_numbering,
    }
    if heading_numbering:
        profile["headingNumbering"] = heading_numbering
    if paragraph_rules:
        profile["paragraphRules"] = paragraph_rules
    if replacements:
        profile["replacements"] = replacements
    if field_rules:
        profile["fieldRules"] = field_rules
    ensure_profile(profile)
    return {
        "ok": True,
        "input": str(path.resolve()),
        "summary": model.get("summary"),
        "observed": {
            "sections": len(sections),
            "styles": len(report.get("styles") or {}),
            "headers": (report.get("marginals") or {}).get("headers"),
            "footers": (report.get("marginals") or {}).get("footers"),
            "placeholders": len(replacements),
            "headingNumbering": bool(heading_numbering),
            "paragraphRules": len(paragraph_rules),
        },
        "profile": profile,
        "unresolved": unresolved,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Analyze a DOCX template into a draft profile.")
    parser.add_argument("--input", required=True, type=Path, help="Official template .docx")
    parser.add_argument("--output", type=Path, help="Write analyzer JSON")
    parser.add_argument("--profile-out", type=Path, help="Write draft profile JSON")
    args = parser.parse_args()
    try:
        if not args.input.is_file():
            raise FileNotFoundError(args.input)
        result = analyze_template(args.input)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        if args.profile_out:
            args.profile_out.parent.mkdir(parents=True, exist_ok=True)
            args.profile_out.write_text(json.dumps(result["profile"], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(
            json.dumps(
                {
                    "ok": True,
                    "summary": result.get("summary"),
                    "output": str(args.output.resolve()) if args.output else None,
                    "profile": str(args.profile_out.resolve()) if args.profile_out else None,
                },
                ensure_ascii=False,
            )
        )
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
