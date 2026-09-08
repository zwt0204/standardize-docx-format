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


def cm_to_inches(value: float) -> float:
    return round(value / 2.54, 4)


def mm_to_inches(value: float) -> float:
    return round(value / 25.4, 4)


def read_source(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".txt", ".md", ".html", ".htm", ".json"}:
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
    validation: dict[str, Any] = {
        "requiredTexts": ["摘要", "Abstract", "目录", "参考文献"],
        "noPlaceholders": ["\\{\\{[^}]+\\}\\}"],
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
            add_rule(rules, "pageNumbers", "front", "upper-roman", line, "medium")
        if re.search(r"正文.{0,8}页码.{0,8}[1一]|阿拉伯数字", line):
            add_rule(rules, "pageNumbers", "body", "decimal", line, "medium")
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

    profile = {
        "name": f"从规范编译: {source_name}",
        "requirements": {
            "sourceType": "description",
            "sources": [source_name],
            "unresolved": unresolved,
            "conflicts": conflicts,
        },
        "document": {
            "type": "thesis",
            "sections": DEFAULT_THESIS_SECTIONS,
        },
        "page": page or {"size": "a4", "orientation": "portrait"},
        "defaultRun": default_run,
        "styles": styles,
        "fields": {"updateOnOpen": True},
        "validation": validation,
    }
    if heading_numbering:
        profile["headingNumbering"] = heading_numbering

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
