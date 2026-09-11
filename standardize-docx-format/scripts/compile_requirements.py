#!/usr/bin/env python3
"""Compile formatting requirements from text/PDF-extracted text into a Requirement IR and draft profile."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

from profile_schema import ensure_profile


CHINESE_SIZE_PT = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
}

FONT_ALIASES = {
    "宋体": "宋体",
    "新宋体": "宋体",
    "黑体": "黑体",
    "楷体": "楷体",
    "楷体_gb2312": "楷体",
    "仿宋": "仿宋",
    "仿宋_gb2312": "仿宋",
    "微软雅黑": "微软雅黑",
    "times new roman": "Times New Roman",
    "times": "Times New Roman",
    "arial": "Arial",
}

DEFAULT_THESIS_SECTIONS = [
    {"id": "cover", "required": True},
    {"id": "abstract", "required": True, "languages": ["zh", "en"]},
    {"id": "toc", "required": True},
    {"id": "chapter", "required": True, "repeatable": True, "min": 1},
    {"id": "references", "required": True},
]

DOCUMENT_TYPE_HINTS = [
    ("technical-report", re.compile(r"技术报告|technical report", re.I)),
    ("course-paper", re.compile(r"课程论文|课程设计|course paper", re.I)),
    ("journal", re.compile(r"期刊|journal article", re.I)),
    ("thesis", re.compile(r"毕业论文|学位论文|毕业设计|thesis|dissertation", re.I)),
]


def cm_to_inches(value: float) -> float:
    return round(value / 2.54, 4)


def mm_to_inches(value: float) -> float:
    return round(value / 25.4, 4)


def strip_html(text: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p>", "\n", text)
    text = re.sub(r"(?i)</h[1-6]>", "\n", text)
    text = re.sub(r"(?i)</li>", "\n", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    return text


def infer_document_type(text: str) -> str:
    for type_name, pattern in DOCUMENT_TYPE_HINTS:
        if pattern.search(text):
            return type_name
    return "thesis"


def document_sections_for(doc_type: str) -> list[dict[str, Any]]:
    if doc_type == "thesis":
        return list(DEFAULT_THESIS_SECTIONS)
    if doc_type == "course-paper":
        return [
            {"id": "cover", "required": False},
            {"id": "abstract", "required": False, "languages": ["zh"]},
            {"id": "toc", "required": False},
            {"id": "chapter", "required": True, "repeatable": True, "min": 1},
            {"id": "references", "required": True},
        ]
    if doc_type == "technical-report":
        return [
            {"id": "cover", "required": True},
            {"id": "abstract", "required": False, "languages": ["zh", "en"]},
            {"id": "toc", "required": True},
            {"id": "chapter", "required": True, "repeatable": True, "min": 1},
            {"id": "references", "required": True},
            {"id": "appendix", "required": False},
        ]
    return list(DEFAULT_THESIS_SECTIONS)


def read_source(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".html", ".htm"}:
        return strip_html(path.read_text(encoding="utf-8"))
    if suffix in {".txt", ".md", ".json"}:
        return path.read_text(encoding="utf-8")
    if suffix == ".pdf":
        pdftotext = shutil.which("pdftotext")
        if pdftotext:
            result = subprocess.run(
                [pdftotext, "-layout", str(path), "-"],
                check=False,
                capture_output=True,
                text=True,
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout
        raise ValueError("PDF extraction needs pdftotext, or pass a .txt dump of the specification")
    return path.read_text(encoding="utf-8")


def add_rule(rules: list[dict[str, Any]], target: str, property_name: str, value: Any, evidence: str, confidence: str = "high") -> None:
    rules.append(
        {
            "target": target,
            "property": property_name,
            "value": value,
            "evidence": evidence.strip()[:160],
            "confidence": confidence,
        }
    )


def find_font(text: str) -> str | None:
    lower = text.lower()
    for alias, canonical in FONT_ALIASES.items():
        if alias in text or alias in lower:
            return canonical
    return None


def find_size_pt(text: str) -> float | None:
    for name, size in CHINESE_SIZE_PT.items():
        if name in text:
            return size
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:pt|磅|磅字)", text, flags=re.I)
    if match:
        return float(match.group(1))
    return None


def compile_text(source_text: str, source_name: str = "specification") -> dict[str, Any]:
    rules: list[dict[str, Any]] = []
    unresolved: list[str] = []
    conflicts: list[str] = []
    lines = [re.sub(r"\s+", " ", line).strip() for line in source_text.splitlines() if line.strip()]

    page: dict[str, Any] = {}
    default_run: dict[str, Any] = {"latinFont": "Times New Roman", "complexScriptFont": "Times New Roman"}
    styles: dict[str, Any] = {}
    heading_numbering: dict[str, Any] = {}
    page_numbers: list[dict[str, Any]] = []
    headers_footers: list[dict[str, Any]] = []
    caption_numbering: dict[str, Any] = {}
    document_type = infer_document_type(source_text)
    validation: dict[str, Any] = {
        "requiredTexts": ["摘要", "Abstract", "目录", "参考文献"] if document_type == "thesis" else [],
        "noPlaceholders": ["\\{\\{[^}]+\\}\\}"],
        "requiredFields": {},
    }

    def style_bucket(style_id: str) -> dict[str, Any]:
        return styles.setdefault(style_id, {"run": {}, "paragraph": {}})

    for line in lines:
        if re.search(r"\bA4\b|A4纸|纸型.{0,4}A4", line, flags=re.I):
            page["size"] = "a4"
            page["orientation"] = "portrait"
            add_rule(rules, "page", "size", "a4", line)
        if re.search(r"纵向|portrait", line, flags=re.I):
            page["orientation"] = "portrait"
            add_rule(rules, "page", "orientation", "portrait", line)
        margin_matches = list(
            re.finditer(
                r"(上|下|左|右)[边頁页]?[距边]?\s*[为是:]?\s*(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)",
                line,
                flags=re.I,
            )
        )
        for margin_match in margin_matches:
            side_map = {"上": "top", "下": "bottom", "左": "left", "右": "right"}
            side = side_map[margin_match.group(1)]
            amount = float(margin_match.group(2))
            unit = margin_match.group(3).lower()
            inches = cm_to_inches(amount) if unit in {"cm", "厘米"} else mm_to_inches(amount) if unit in {"mm", "毫米"} else amount
            page.setdefault("marginsIn", {})[side] = inches
            add_rule(rules, "page.margins", side, inches, line)

        if "页边距" in line and not margin_matches:
            unresolved.append(f"页边距未量化: {line}")

        if re.search(r"正文|主体|body", line) and (find_font(line) or find_size_pt(line) or "行距" in line or "倍" in line):
            font = find_font(line)
            size = find_size_pt(line)
            if font:
                if font in {"Times New Roman", "Arial"}:
                    default_run["latinFont"] = font
                    add_rule(rules, "body", "latinFont", font, line)
                else:
                    default_run["eastAsiaFont"] = font
                    add_rule(rules, "body", "eastAsiaFont", font, line)
            if size:
                default_run["sizePt"] = size
                add_rule(rules, "body", "sizePt", size, line)
            spacing = re.search(r"(?:行距|行间距)\s*[为是:]?\s*(\d+(?:\.\d+)?)\s*倍", line)
            if spacing:
                style_bucket("Normal")["paragraph"]["lineSpacing"] = float(spacing.group(1))
                add_rule(rules, "body", "lineSpacing", float(spacing.group(1)), line)
            if "两端对齐" in line or "justify" in line.lower():
                style_bucket("Normal")["paragraph"]["alignment"] = "justify"
                add_rule(rules, "body", "alignment", "justify", line)
            indent = re.search(r"(?:首行缩进)\s*(\d+(?:\.\d+)?)\s*字符", line)
            if indent:
                style_bucket("Normal")["paragraph"]["firstLineChars"] = float(indent.group(1))
                add_rule(rules, "body", "firstLineChars", float(indent.group(1)), line)
            elif "首行缩进" in line and "两" in line:
                style_bucket("Normal")["paragraph"]["firstLineChars"] = 2
                add_rule(rules, "body", "firstLineChars", 2, line)

        heading_level = None
        if re.search(r"一级标题|章标题|Heading\s*1", line, flags=re.I):
            heading_level = "Heading1"
        elif re.search(r"二级标题|节标题|Heading\s*2", line, flags=re.I):
            heading_level = "Heading2"
        elif re.search(r"三级标题|Heading\s*3", line, flags=re.I):
            heading_level = "Heading3"
        if heading_level:
            font = find_font(line) or "黑体"
            size = find_size_pt(line)
            bucket = style_bucket(heading_level)
            bucket["run"]["eastAsiaFont"] = font
            bucket["run"]["latinFont"] = "Times New Roman"
            bucket["run"]["bold"] = True
            bucket["paragraph"]["keepNext"] = True
            if heading_level == "Heading1":
                bucket["paragraph"]["outlineLevel"] = 0
            elif heading_level == "Heading2":
                bucket["paragraph"]["outlineLevel"] = 1
            else:
                bucket["paragraph"]["outlineLevel"] = 2
            if size:
                bucket["run"]["sizePt"] = size
            add_rule(rules, heading_level, "font", font, line)
            if size:
                add_rule(rules, heading_level, "sizePt", size, line)

        if "罗马" in line and "页码" in line:
            add_rule(rules, "pageNumbers.front", "format", "upper-roman", line, "medium")
            page_numbers.append({"section": 0, "format": "upper-roman", "start": 1})
            unresolved.append("前置罗马页码已写入草稿 section 0；请按实际分节索引调整 pageNumbers")
        if re.search(r"正文.{0,8}页码.{0,8}[1一]|页码.{0,6}阿拉伯", line):
            add_rule(rules, "pageNumbers.body", "format", "decimal", line, "medium")
            page_numbers.append({"section": 1, "format": "decimal", "start": 1})
            unresolved.append("正文阿拉伯页码已写入草稿 section 1；请确认正文起点所在节")
        if re.search(r"一、|（一）|1\.", line) and "标题" in line:
            heading_numbering = {
                "enabled": True,
                "styles": ["Heading1", "Heading2", "Heading3"],
                "patterns": ["%1、", "（%2）", "%3."],
                "formats": ["chinese-counting", "chinese-counting", "decimal"],
                "starts": [1, 1, 1],
                "suffix": "space",
            }
            add_rule(rules, "headingNumbering", "enabled", True, line, "medium")

        if re.search(r"页眉", line):
            add_rule(rules, "headersFooters.header", "present", True, line, "medium")
            text = None
            quoted = re.search(r"[“\"']([^”\"']+)[”\"']", line)
            if quoted:
                text = quoted.group(1)
            headers_footers.append(
                {
                    "section": 1,
                    "kind": "header",
                    "type": "even" if ("偶" in line) else "first" if ("首页" in line or "封面" in line) else "default",
                    "action": "set" if text else "inherit",
                    "paragraphs": [{"alignment": "center", "segments": [{"text": text}]}] if text else [],
                    "note": line[:160],
                }
            )
            if "奇偶" in line or "偶数" in line:
                unresolved.append("规范要求奇偶页不同页眉；草稿只记录 inherit/set，请按官方模板确认内容")
            if not text:
                unresolved.append(f"页眉内容未量化: {line}")
        if re.search(r"页脚", line) or re.search(r"页码.{0,6}(居中|底部|页脚)", line):
            add_rule(rules, "headersFooters.footer", "page", True, line, "medium")
            headers_footers.append(
                {
                    "section": 1,
                    "kind": "footer",
                    "type": "default",
                    "action": "set",
                    "paragraphs": [
                        {
                            "alignment": "center",
                            "segments": [{"field": "PAGE", "result": "1"}],
                        }
                    ],
                }
            )
            validation.setdefault("requiredFields", {})["PAGE"] = 1

        if re.search(r"目录", line) and re.search(r"自动|域|字段|三级|两级|1-2|1-3", line):
            add_rule(rules, "fields", "TOC", True, line, "medium")
            validation.setdefault("requiredFields", {})["TOC"] = 1
        if re.search(r"题注|图题|表题", line):
            add_rule(rules, "captions", "present", True, line, "medium")
            if re.search(r"阿拉伯|1-1|章号", line) and (
                heading_numbering.get("formats") and str(heading_numbering["formats"][0]).startswith("chinese")
            ):
                caption_numbering = {
                    "chapterStyle": "Heading 1",
                    "headingDisplay": "chinese-counting",
                    "chapterDisplay": "decimal",
                    "separator": "-",
                    "strategy": "needs-confirmation",
                }
                unresolved.append("正文标题可能使用中文序号，图表要求阿拉伯章号；请确认 captionNumbering.strategy")
            elif re.search(r"图\s*\d|表\s*\d|按章", line):
                caption_numbering = {
                    "chapterStyle": "Heading 1",
                    "separator": "-",
                    "strategy": "styleref-as-displayed",
                }

        if re.search(r"参考文献", line) and re.search(r"著者-出版年|顺序编码|GB/T\s*7714", line):
            unresolved.append(f"参考文献体例需确认: {line}")

        if re.search(r"适当|美观|大致|左右|参考学校", line):
            unresolved.append(f"不可量化: {line}")

    if "eastAsiaFont" not in default_run:
        default_run["eastAsiaFont"] = "宋体"
        unresolved.append("正文中文字体未明确，草稿默认宋体")
    if "sizePt" not in default_run:
        default_run["sizePt"] = 12
        unresolved.append("正文字号未明确，草稿默认小四/12pt")
    default_run["updateTheme"] = True

    normal = style_bucket("Normal")
    normal["name"] = "正文"
    normal["run"] = {
        "latinFont": default_run.get("latinFont"),
        "eastAsiaFont": default_run.get("eastAsiaFont"),
        "complexScriptFont": default_run.get("complexScriptFont"),
        "sizePt": default_run.get("sizePt"),
    }
    normal.setdefault("paragraph", {}).setdefault("lineSpacing", 1.5)
    normal["paragraph"].setdefault("alignment", "justify")
    normal["paragraph"].setdefault("firstLineChars", 2)
    normal["paragraph"].setdefault("widowControl", True)

    seen: dict[tuple[str, str], Any] = {}
    for rule in rules:
        key = (rule["target"], rule["property"])
        if key in seen and seen[key] != rule["value"]:
            conflicts.append(f"{rule['target']}.{rule['property']}: {seen[key]} vs {rule['value']}")
        seen[key] = rule["value"]

    if not validation.get("requiredFields"):
        validation.pop("requiredFields", None)
    if not validation.get("requiredTexts"):
        validation.pop("requiredTexts", None)

    unique_page_numbers = []
    seen_sections: set[int] = set()
    for item in page_numbers:
        section = item.get("section")
        if section in seen_sections:
            continue
        seen_sections.add(section)
        unique_page_numbers.append(item)

    unique_headers = []
    seen_hf: set[tuple[Any, Any, Any]] = set()
    for item in headers_footers:
        key = (item.get("section"), item.get("kind"), item.get("type"))
        if key in seen_hf:
            continue
        seen_hf.add(key)
        unique_headers.append(item)

    profile = {
        "name": f"从规范编译: {source_name}",
        "profileVersion": "2.1",
        "requirements": {
            "sourceType": "description",
            "sources": [source_name],
            "unresolved": unresolved,
            "conflicts": conflicts,
        },
        "document": {
            "type": document_type,
            "sections": document_sections_for(document_type),
        },
        "page": page or {"size": "a4", "orientation": "portrait"},
        "defaultRun": default_run,
        "styles": styles,
        "fields": {"updateOnOpen": True},
        "validation": validation,
    }
    if heading_numbering:
        profile["headingNumbering"] = heading_numbering
    if unique_page_numbers:
        profile["pageNumbers"] = unique_page_numbers
    if unique_headers:
        profile["headersFooters"] = unique_headers
    if caption_numbering:
        profile["captionNumbering"] = caption_numbering
    ensure_profile(profile)

    return {
        "ok": not conflicts,
        "source": source_name,
        "rules": rules,
        "unresolved": unresolved,
        "conflicts": conflicts,
        "profile": profile,
        "notes": [
            "This compiler extracts measurable clauses; Codex should still review the draft profile.",
            "Do not apply while unresolved or conflicts remain unless the user accepts --allow-unresolved.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Compile a draft profile from a formatting specification.")
    parser.add_argument("--input", required=True, type=Path, help="Specification .txt/.md/.html or PDF")
    parser.add_argument("--output", type=Path, help="Write Requirement IR JSON")
    parser.add_argument("--profile-out", type=Path, help="Write the draft profile JSON")
    args = parser.parse_args()
    try:
        text = read_source(args.input)
        result = compile_text(text, args.input.name)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        if args.profile_out:
            args.profile_out.parent.mkdir(parents=True, exist_ok=True)
            args.profile_out.write_text(json.dumps(result["profile"], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(
            json.dumps(
                {
                    "ok": result["ok"],
                    "rules": len(result["rules"]),
                    "unresolved": len(result["unresolved"]),
                    "conflicts": len(result["conflicts"]),
                    "output": str(args.output.resolve()) if args.output else None,
                    "profile": str(args.profile_out.resolve()) if args.profile_out else None,
                },
                ensure_ascii=False,
            )
        )
        return 0 if result["ok"] else 2
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
