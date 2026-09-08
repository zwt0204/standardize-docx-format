#!/usr/bin/env python3
"""Compare a DOCX against a profile and emit an explainable formatting diff."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from audit_docx import audit
from document_model import build_model, document_spec


PAGE_SIZES_IN = {
    "a3": (11.693, 16.535),
    "a4": (8.268, 11.693),
    "a5": (5.827, 8.268),
    "letter": (8.5, 11.0),
    "legal": (8.5, 14.0),
    "tabloid": (11.0, 17.0),
}

FORMAT_MAP = {
    "decimal": "decimal",
    "lower-roman": "lowerRoman",
    "upper-roman": "upperRoman",
    "lower-alpha": "lowerLetter",
    "upper-alpha": "upperLetter",
}

REGION_IDS = {
    "cover": "hasCover",
    "abstract": "hasAbstractZh",
    "abstract_zh": "hasAbstractZh",
    "abstract_en": "hasAbstractEn",
    "toc": "hasToc",
    "chapter": "chapterCount",
    "references": "hasReferences",
    "appendix": "hasAppendix",
}


def approx(actual: Any, expected: Any, tolerance: float = 0.04) -> bool:
    if actual is None or expected is None:
        return actual == expected
    try:
        return abs(float(actual) - float(expected)) <= tolerance
    except (TypeError, ValueError):
        return actual == expected


def add_finding(
    findings: list[dict[str, Any]],
    category: str,
    name: str,
    ok: bool,
    expected: Any,
    actual: Any,
    severity: str = "error",
    affected: int | None = None,
    suggestion: str | None = None,
) -> None:
    findings.append(
        {
            "category": category,
            "name": name,
            "ok": ok,
            "severity": "info" if ok else severity,
            "expected": expected,
            "actual": actual,
            "affected": affected,
            "suggestion": suggestion,
        }
    )


def expected_page_size(page: dict[str, Any]) -> tuple[float | None, float | None]:
    if "widthIn" in page or "heightIn" in page:
        return page.get("widthIn"), page.get("heightIn")
    size = str(page.get("size") or "").lower()
    if size in PAGE_SIZES_IN:
        width, height = PAGE_SIZES_IN[size]
        if str(page.get("orientation") or "portrait") == "landscape":
            return height, width
        return width, height
    return None, None


def first_line_inches(paragraph: dict[str, Any], size_pt: Any) -> float | None:
    if "firstLineIn" in paragraph:
        return float(paragraph["firstLineIn"])
    if "firstLineChars" in paragraph and size_pt:
        return round(float(paragraph["firstLineChars"]) * float(size_pt) * 20 / 1440, 4)
    return None


def style_usage(report: dict[str, Any], style_id: str) -> int:
    usage = (report.get("content") or {}).get("styleUsage") or {}
    return int(usage.get(style_id, 0))


def diff_page(findings: list[dict[str, Any]], report: dict[str, Any], profile: dict[str, Any]) -> None:
    page = profile.get("page") or {}
    if not isinstance(page, dict) or not page:
        return
    sections = report.get("sections") or []
    if not sections:
        add_finding(findings, "page", "sections", False, ">=1", 0, suggestion="Document has no sectPr")
        return
    expected_width, expected_height = expected_page_size(page)
    orientation = page.get("orientation")
    margins = page.get("marginsIn") or {}
    for section in sections:
        actual_page = section.get("page") or {}
        index = section.get("index")
        if expected_width is not None:
            add_finding(
                findings,
                "page",
                f"section[{index}].widthIn",
                approx(actual_page.get("widthIn"), expected_width, 0.05),
                expected_width,
                actual_page.get("widthIn"),
            )
        if expected_height is not None:
            add_finding(
                findings,
                "page",
                f"section[{index}].heightIn",
                approx(actual_page.get("heightIn"), expected_height, 0.05),
                expected_height,
                actual_page.get("heightIn"),
            )
        if orientation:
            actual_orientation = str(actual_page.get("orientation") or "").replace("/default", "")
            add_finding(
                findings,
                "page",
                f"section[{index}].orientation",
                actual_orientation == orientation or (orientation == "portrait" and actual_orientation in {"portrait", ""}),
                orientation,
                actual_page.get("orientation"),
            )
        actual_margins = actual_page.get("marginsIn") or {}
        for key, expected in margins.items():
            add_finding(
                findings,
                "page",
                f"section[{index}].margins.{key}",
                approx(actual_margins.get(key), expected, 0.05),
                expected,
                actual_margins.get(key),
            )


def diff_run(findings: list[dict[str, Any]], actual: dict[str, Any] | None, expected: dict[str, Any], prefix: str, affected: int | None) -> None:
    actual = actual or {}
    mapping = {
        "latinFont": "latin",
        "eastAsiaFont": "eastAsia",
        "complexScriptFont": "complexScript",
        "sizePt": "sizePt",
        "bold": "bold",
        "italic": "italic",
        "color": "color",
    }
    for expected_key, actual_key in mapping.items():
        if expected_key not in expected:
            continue
        want = expected[expected_key]
        got = actual.get(actual_key)
        ok = approx(got, want, 0.25) if expected_key == "sizePt" else got == want
        add_finding(
            findings,
            "run",
            f"{prefix}.{expected_key}",
            ok,
            want,
            got,
            affected=affected,
            suggestion=f"set {prefix} {expected_key} to {want}",
        )


def diff_paragraph(
    findings: list[dict[str, Any]],
    actual: dict[str, Any] | None,
    expected: dict[str, Any],
    prefix: str,
    size_pt: Any,
    affected: int | None,
) -> None:
    actual = actual or {}
    for key in (
        "alignment",
        "lineSpacing",
        "lineSpacingExactPt",
        "spaceBeforePt",
        "spaceAfterPt",
        "keepNext",
        "keepLines",
        "pageBreakBefore",
        "widowControl",
        "outlineLevel",
    ):
        if key not in expected:
            continue
        want = expected[key]
        got = actual.get(key)
        tolerance = 0.06 if key in {"lineSpacing", "spaceBeforePt", "spaceAfterPt", "lineSpacingExactPt"} else 0.0
        ok = approx(got, want, tolerance) if tolerance else got == want
        add_finding(
            findings,
            "paragraph",
            f"{prefix}.{key}",
            ok,
            want,
            got,
            affected=affected,
            suggestion=f"set {prefix} {key} to {want}",
        )
    expected_first = first_line_inches(expected, size_pt)
    if expected_first is not None:
        actual_first = actual.get("firstLineIn")
        add_finding(
            findings,
            "paragraph",
            f"{prefix}.firstLineIn",
            approx(actual_first, expected_first, 0.04),
            expected_first,
            actual_first,
            affected=affected,
            suggestion=f"set {prefix} first-line indent to {expected.get('firstLineChars', expected_first)}",
        )


def diff_styles(findings: list[dict[str, Any]], report: dict[str, Any], profile: dict[str, Any]) -> None:
    default_run = profile.get("defaultRun") or {}
    if isinstance(default_run, dict) and default_run:
        usage = int((report.get("content") or {}).get("paragraphs") or 0)
        diff_run(findings, report.get("defaultRun"), default_run, "defaultRun", usage)
    styles = profile.get("styles") or {}
    catalog = report.get("styles") or {}
    if not isinstance(styles, dict):
        return
    for style_id, config in styles.items():
        if not isinstance(config, dict):
            continue
        actual = catalog.get(style_id)
        affected = style_usage(report, style_id)
        add_finding(
            findings,
            "style",
            f"styles.{style_id}.exists",
            actual is not None,
            True,
            actual is not None,
            affected=affected,
            suggestion=f"create style {style_id}",
        )
        if actual is None:
            continue
        run_config = config.get("run") or {}
        paragraph_config = config.get("paragraph") or {}
        size_pt = (run_config or {}).get("sizePt") or default_run.get("sizePt")
        if run_config:
            merged_run = dict(actual.get("run") or {})
            merged_run.setdefault("latin", (actual.get("fonts") or {}).get("latin"))
            merged_run.setdefault("eastAsia", (actual.get("fonts") or {}).get("eastAsia"))
            merged_run.setdefault("sizePt", actual.get("sizePt"))
            diff_run(findings, merged_run, run_config, f"styles.{style_id}", affected)
        if paragraph_config:
            diff_paragraph(
                findings,
                actual.get("paragraph") or {},
                paragraph_config,
                f"styles.{style_id}",
                size_pt,
                affected,
            )


def diff_page_numbers(findings: list[dict[str, Any]], report: dict[str, Any], profile: dict[str, Any]) -> None:
    sections = report.get("sections") or []
    for item in profile.get("pageNumbers") or []:
        if not isinstance(item, dict) or not isinstance(item.get("section"), int):
            continue
        index = item["section"]
        actual = sections[index].get("pageNumber") if 0 <= index < len(sections) else None
        expected_format = FORMAT_MAP.get(item.get("format"), item.get("format"))
        ok = actual is not None
        if ok and expected_format:
            ok = actual.get("format") == expected_format
        if ok and item.get("start") is not None:
            ok = actual.get("start") == item.get("start")
        add_finding(
            findings,
            "pageNumbers",
            f"pageNumbers.section[{index}]",
            bool(ok),
            {"format": expected_format, "start": item.get("start")},
            actual,
            suggestion="configure w:pgNumType for this section",
        )


def diff_headers(findings: list[dict[str, Any]], report: dict[str, Any], profile: dict[str, Any]) -> None:
    sections = report.get("sections") or []
    for index, item in enumerate(profile.get("headersFooters") or []):
        if not isinstance(item, dict) or item.get("action", "set") == "inherit":
            continue
        section_index = item.get("section")
        kind = item.get("kind")
        reference_type = item.get("type", "default")
        actual_refs = []
        if isinstance(section_index, int) and 0 <= section_index < len(sections):
            actual_refs = sections[section_index].get(f"{kind}References") or []
        ok = any(reference.get("type") == reference_type for reference in actual_refs)
        add_finding(
            findings,
            "headersFooters",
            f"headersFooters[{index}]",
            ok,
            {"section": section_index, "kind": kind, "type": reference_type},
            actual_refs,
            suggestion=f"create {kind} {reference_type} for section {section_index}",
        )


def diff_structure(findings: list[dict[str, Any]], model: dict[str, Any], profile: dict[str, Any]) -> None:
    spec = document_spec(profile)
    summary = model.get("summary") or {}
    for index, item in enumerate(spec.get("sections") or []):
        if not isinstance(item, dict) or not item.get("id"):
            continue
        section_id = str(item["id"])
        required = bool(item.get("required", False))
        key = REGION_IDS.get(section_id)
        if key == "chapterCount":
            actual = int(summary.get("chapterCount") or 0)
            minimum = int(item.get("min", 1 if required else 0))
            ok = actual >= minimum
            add_finding(
                findings,
                "structure",
                f"document.sections[{index}].{section_id}",
                ok,
                {"min": minimum, "repeatable": bool(item.get("repeatable", True))},
                actual,
                severity="error" if required else "warning",
                suggestion="map chapter headings with paragraphRules or Heading1",
            )
            continue
        if key:
            actual = bool(summary.get(key))
            add_finding(
                findings,
                "structure",
                f"document.sections[{index}].{section_id}",
                actual if required else True,
                required,
                actual,
                severity="error" if required and not actual else "info",
                suggestion=f"ensure a recognizable {section_id} region exists",
            )
        languages = item.get("languages") or []
        if section_id in {"abstract", "abstract_zh"} and "en" in languages:
            add_finding(
                findings,
                "structure",
                "document.abstract_en",
                bool(summary.get("hasAbstractEn")),
                True,
                bool(summary.get("hasAbstractEn")),
                severity="warning",
                suggestion="ensure an English Abstract heading exists",
            )


def diff_document(input_path: Path, profile: dict[str, Any]) -> dict[str, Any]:
    report = audit(input_path)
    model = build_model(input_path, profile)
    findings: list[dict[str, Any]] = []
    diff_page(findings, report, profile)
    diff_styles(findings, report, profile)
    diff_page_numbers(findings, report, profile)
    diff_headers(findings, report, profile)
    diff_structure(findings, model, profile)
    failed = [item for item in findings if not item["ok"]]
    categories: dict[str, dict[str, int]] = {}
    for item in findings:
        bucket = categories.setdefault(item["category"], {"total": 0, "failed": 0})
        bucket["total"] += 1
        if not item["ok"]:
            bucket["failed"] += 1
    scores = {}
    for category, bucket in categories.items():
        total = max(bucket["total"], 1)
        scores[category] = round(100 * (total - bucket["failed"]) / total, 1)
    overall = round(sum(scores.values()) / max(len(scores), 1), 1) if scores else 100.0
    return {
        "ok": not failed,
        "input": str(input_path.resolve()),
        "profile": profile.get("name"),
        "compliancePercent": overall,
        "categoryScores": scores,
        "findings": findings,
        "failed": failed,
        "summary": model.get("summary"),
    }


def render_text(diff: dict[str, Any]) -> str:
    lines = [
        "Document Diff",
        "────────────────────────",
        f"Compliance: {diff.get('compliancePercent')}%",
        "",
    ]
    current = None
    for item in diff.get("findings", []):
        if item["category"] != current:
            current = item["category"]
            lines.append(current)
        mark = "✓" if item["ok"] else "✗"
        extra = ""
        if not item["ok"]:
            extra = f"      expected: {item['expected']}\n      actual:   {item['actual']}"
            if item.get("affected"):
                extra += f"\n      affected: {item['affected']}"
        lines.append(f"  {mark} {item['name']}")
        if extra:
            lines.append(extra)
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Diff a DOCX against a formatting profile.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--profile", required=True, type=Path)
    parser.add_argument("--output", type=Path, help="Write JSON diff")
    parser.add_argument("--text", action="store_true", help="Print a human-readable summary")
    args = parser.parse_args()
    try:
        profile = json.loads(args.profile.read_text(encoding="utf-8"))
        if not isinstance(profile, dict):
            raise ValueError("Profile root must be an object")
        result = diff_document(args.input, profile)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        if args.text or not args.output:
            print(render_text(result), end="")
        else:
            print(json.dumps({"ok": result["ok"], "compliancePercent": result["compliancePercent"], "output": str(args.output.resolve())}, ensure_ascii=False))
        return 0 if result["ok"] else 2
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
