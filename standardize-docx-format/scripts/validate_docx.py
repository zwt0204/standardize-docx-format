#!/usr/bin/env python3
"""Validate a DOCX against the measurable rules in a standardization profile."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any
from xml.dom import Node, minidom

from audit_docx import audit


def node_text(node: Node) -> str:
    return "".join(child.data for child in node.childNodes if child.nodeType == Node.TEXT_NODE)


def direct_paragraph_text_nodes(paragraph: Node) -> list[Node]:
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


def paragraph_texts(document: minidom.Document) -> list[str]:
    out = []
    for paragraph in document.getElementsByTagName("w:p"):
        text = "".join(node_text(node) for node in direct_paragraph_text_nodes(paragraph))
        out.append(text)
    return out


def metric_count(text: str, mode: str) -> int:
    if mode == "characters":
        return len(text)
    if mode == "charactersNoWhitespace":
        return len(re.sub(r"\s+", "", text))
    if mode == "cjkCharacters":
        return len(re.findall(r"[\u3400-\u4dbf\u4e00-\u9fff]", text))
    if mode in {"latinWords", "words"}:
        latin = len(re.findall(r"[A-Za-z]+(?:['’-][A-Za-z]+)*", text))
        if mode == "latinWords":
            return latin
        return latin + len(re.findall(r"[\u3400-\u4dbf\u4e00-\u9fff]", text))
    raise ValueError(f"Unsupported word-count mode: {mode}")


def add_check(
    checks: list[dict[str, Any]], name: str, ok: bool, expected: Any = None, actual: Any = None
) -> None:
    checks.append({"name": name, "ok": bool(ok), "expected": expected, "actual": actual})


def count_matches(text: str, pattern: str, regex: bool) -> int:
    return len(re.findall(pattern, text, flags=re.MULTILINE)) if regex else text.count(pattern)


def validate_text_rules(
    checks: list[dict[str, Any]], text: str, rules: list[Any], required: bool
) -> None:
    label = "requiredTexts" if required else "forbiddenTexts"
    for index, rule in enumerate(rules):
        if isinstance(rule, str):
            pattern, regex, minimum, maximum = rule, False, (1 if required else 0), (None if required else 0)
        elif isinstance(rule, dict) and "pattern" in rule:
            pattern = str(rule["pattern"])
            regex = bool(rule.get("regex", False))
            minimum = int(rule.get("min", 1 if required else 0))
            maximum = rule.get("max", None if required else 0)
            maximum = int(maximum) if maximum is not None else None
        else:
            raise ValueError(f"{label}[{index}] must be a string or object with pattern")
        actual = count_matches(text, pattern, regex)
        ok = actual >= minimum and (maximum is None or actual <= maximum)
        add_check(checks, f"{label}[{index}]", ok, {"min": minimum, "max": maximum}, actual)


def find_marker(paragraphs: list[str], pattern: str, start: int = 0) -> int | None:
    compiled = re.compile(pattern)
    for index in range(start, len(paragraphs)):
        if compiled.search(paragraphs[index]):
            return index
    return None


def validate_word_counts(
    checks: list[dict[str, Any]], paragraphs: list[str], rules: list[Any]
) -> None:
    for index, rule in enumerate(rules):
        if not isinstance(rule, dict) or "startRegex" not in rule or "endRegex" not in rule:
            raise ValueError("Each validation.wordCounts item needs startRegex and endRegex")
        start = find_marker(paragraphs, str(rule["startRegex"]))
        end = find_marker(paragraphs, str(rule["endRegex"]), (start + 1) if start is not None else 0)
        name = str(rule.get("name", f"wordCounts[{index}]"))
        if start is None or end is None or end <= start:
            add_check(checks, name, False, "start/end markers", {"start": start, "end": end})
            continue
        include_markers = bool(rule.get("includeMarkers", False))
        selected = paragraphs[start : end + 1] if include_markers else paragraphs[start + 1 : end]
        value = "\n".join(selected)
        mode = str(rule.get("mode", "charactersNoWhitespace"))
        actual = metric_count(value, mode)
        minimum = rule.get("min")
        maximum = rule.get("max")
        ok = (minimum is None or actual >= int(minimum)) and (maximum is None or actual <= int(maximum))
        add_check(checks, name, ok, {"mode": mode, "min": minimum, "max": maximum}, actual)


def validate_profile_expectations(
    checks: list[dict[str, Any]], report: dict[str, Any], profile: dict[str, Any]
) -> None:
    sections = report.get("sections", [])
    for item in profile.get("pageNumbers", []):
        if not isinstance(item, dict) or not isinstance(item.get("section"), int):
            continue
        index = item["section"]
        actual = sections[index].get("pageNumber") if 0 <= index < len(sections) else None
        expected_format = item.get("format")
        format_map = {
            "decimal": "decimal",
            "lower-roman": "lowerRoman",
            "upper-roman": "upperRoman",
            "lower-alpha": "lowerLetter",
            "upper-alpha": "upperLetter",
        }
        expected = {
            "format": format_map.get(expected_format, expected_format),
            "start": item.get("start"),
        }
        ok = actual is not None
        if ok and expected["format"] is not None:
            ok = actual.get("format") == expected["format"]
        if ok and expected["start"] is not None:
            ok = actual.get("start") == expected["start"]
        add_check(checks, f"pageNumbers.section[{index}]", ok, expected, actual)

    for index, item in enumerate(profile.get("headersFooters", [])):
        if not isinstance(item, dict) or item.get("action", "set") == "inherit":
            continue
        section_index = item.get("section")
        kind = item.get("kind")
        reference_type = item.get("type", "default")
        actual_refs = []
        if isinstance(section_index, int) and 0 <= section_index < len(sections):
            actual_refs = sections[section_index].get(f"{kind}References", [])
        ok = any(reference.get("type") == reference_type for reference in actual_refs)
        add_check(
            checks,
            f"headersFooters[{index}]",
            ok,
            {"section": section_index, "kind": kind, "type": reference_type},
            actual_refs,
        )


def validate_docx(input_path: Path, profile_path: Path) -> dict[str, Any]:
    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    if not isinstance(profile, dict):
        raise ValueError("Profile root must be an object")
    validation = profile.get("validation", {})
    if not isinstance(validation, dict):
        raise ValueError("validation must be an object")
    report = audit(input_path)
    with zipfile.ZipFile(input_path, "r") as archive:
        document = minidom.parseString(archive.read("word/document.xml"))
    paragraphs = paragraph_texts(document)
    full_text = "\n".join(paragraphs)
    checks: list[dict[str, Any]] = []

    section_rule = validation.get("sectionCount")
    if isinstance(section_rule, int):
        add_check(checks, "sectionCount", len(report["sections"]) == section_rule, section_rule, len(report["sections"]))
    elif isinstance(section_rule, dict):
        actual = len(report["sections"])
        minimum = section_rule.get("min")
        maximum = section_rule.get("max")
        ok = (minimum is None or actual >= int(minimum)) and (maximum is None or actual <= int(maximum))
        add_check(checks, "sectionCount", ok, section_rule, actual)

    validate_text_rules(checks, full_text, validation.get("requiredTexts", []), True)
    validate_text_rules(checks, full_text, validation.get("forbiddenTexts", []), False)
    validate_word_counts(checks, paragraphs, validation.get("wordCounts", []))

    for index, pattern in enumerate(validation.get("noPlaceholders", [])):
        actual = len(re.findall(str(pattern), full_text, flags=re.MULTILINE))
        add_check(checks, f"noPlaceholders[{index}]", actual == 0, 0, actual)

    fields = report.get("fields", {}).get("kinds", {})
    for kind, minimum in validation.get("requiredFields", {}).items():
        actual = int(fields.get(str(kind).upper(), 0))
        add_check(checks, f"requiredFields.{kind}", actual >= int(minimum), int(minimum), actual)

    bookmark_names = set(report.get("bookmarks", {}).get("names", []))
    for name in validation.get("requiredBookmarks", []):
        add_check(checks, f"requiredBookmarks.{name}", str(name) in bookmark_names, True, str(name) in bookmark_names)

    style_usage = report.get("content", {}).get("styleUsage", {})
    for style in validation.get("requiredStylesUsed", []):
        actual = int(style_usage.get(str(style), 0))
        add_check(checks, f"requiredStylesUsed.{style}", actual > 0, ">0", actual)

    validate_profile_expectations(checks, report, profile)
    ok = all(check["ok"] for check in checks)
    return {
        "ok": ok,
        "input": str(input_path.resolve()),
        "profile": str(profile_path.resolve()),
        "checks": checks,
        "failed": [check for check in checks if not check["ok"]],
        "audit": report,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a DOCX against a standardization profile.")
    parser.add_argument("--input", required=True, type=Path, help="DOCX to validate")
    parser.add_argument("--profile", required=True, type=Path, help="UTF-8 JSON profile")
    parser.add_argument("--output", type=Path, help="Write the JSON validation report")
    args = parser.parse_args()
    try:
        if not args.input.is_file():
            raise FileNotFoundError(args.input)
        if not args.profile.is_file():
            raise FileNotFoundError(args.profile)
        result = validate_docx(args.input, args.profile)
        rendered = json.dumps(result, ensure_ascii=False, indent=2)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(rendered + "\n", encoding="utf-8")
            print(json.dumps({"ok": result["ok"], "output": str(args.output.resolve())}, ensure_ascii=False))
        else:
            print(rendered)
        return 0 if result["ok"] else 2
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
