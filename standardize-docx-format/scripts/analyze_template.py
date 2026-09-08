#!/usr/bin/env python3
"""Analyze an official DOCX template and emit a draft profile."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from audit_docx import audit
from document_model import build_model
from compile_requirements import DEFAULT_THESIS_SECTIONS
from diff_profile import PAGE_SIZES_IN


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


def analyze_template(path: Path) -> dict[str, Any]:
    report = audit(path)
    model = build_model(path, {"document": {"type": "thesis"}})
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

    profile = {
        "name": f"从模板分析: {path.name}",
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
        "validation": {
            "sectionCount": {"min": max(len(sections), 1)},
            "requiredStylesUsed": [style_id for style_id in ("Normal", "Heading1") if style_id in styles_out],
        },
    }
    return {
        "ok": True,
        "input": str(path.resolve()),
        "summary": model.get("summary"),
        "observed": {
            "sections": len(sections),
            "styles": len(report.get("styles") or {}),
            "headers": (report.get("marginals") or {}).get("headers"),
            "footers": (report.get("marginals") or {}).get("footers"),
        },
        "profile": profile,
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
