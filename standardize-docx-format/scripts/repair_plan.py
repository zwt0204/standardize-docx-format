#!/usr/bin/env python3
"""Turn a profile diff into an explainable, confirmable repair plan."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from diff_profile import diff_document
from visual_qa import visual_qa


ACTION_HINTS = {
    "defaultRun": "set_default_run",
    "styles": "set_style",
    "page": "set_page",
    "pageNumbers": "set_page_number",
    "headersFooters": "set_header_footer",
    "structure": "map_structure",
    "paragraph": "set_paragraph",
    "run": "set_run",
    "style": "set_style",
}


def action_for(finding: dict[str, Any]) -> str:
    name = str(finding.get("name") or "")
    category = str(finding.get("category") or "")
    if name.startswith("defaultRun"):
        return "set_default_run"
    if name.startswith("styles."):
        return "set_style"
    return ACTION_HINTS.get(category, "inspect")


def target_for(finding: dict[str, Any]) -> str:
    name = str(finding.get("name") or "")
    if name.startswith("styles."):
        parts = name.split(".")
        return parts[1] if len(parts) > 1 else name
    if name.startswith("defaultRun"):
        return "defaultRun"
    if "section[" in name:
        return name
    return finding.get("category") or name


def property_for(finding: dict[str, Any]) -> str:
    name = str(finding.get("name") or "")
    return name.rsplit(".", 1)[-1]


VISUAL_ACTIONS = {
    "orphan_heading": "set_keep_next",
    "orphan_heading_risk": "set_keep_next",
    "heading_body_split": "set_keep_next",
    "figure_caption_split": "set_keep_next",
    "table_caption_split": "set_keep_next",
    "image_overflow": "scale_drawing",
    "chapter_midpage": "set_page_break_before",
}


def visual_step(issue: dict[str, Any]) -> dict[str, Any]:
    issue_type = str(issue.get("type") or "visual")
    action = issue.get("repair") or VISUAL_ACTIONS.get(issue_type, "inspect")
    return {
        "action": action,
        "target": f"paragraph[{issue.get('paragraphIndex')}]",
        "property": issue_type,
        "from": issue.get("page"),
        "to": issue.get("suggestion"),
        "affected": 1,
        "severity": issue.get("severity") or "warning",
        "suggestion": issue.get("message") or issue.get("suggestion"),
        "source": issue_type,
        "page": issue.get("page"),
        "paragraphIndex": issue.get("paragraphIndex"),
        "visual": True,
    }


def build_repair_plan(diff: dict[str, Any], visual: dict[str, Any] | None = None) -> dict[str, Any]:
    steps: list[dict[str, Any]] = []
    for finding in diff.get("failed") or []:
        steps.append(
            {
                "action": action_for(finding),
                "target": target_for(finding),
                "property": property_for(finding),
                "from": finding.get("actual"),
                "to": finding.get("expected"),
                "affected": finding.get("affected"),
                "severity": finding.get("severity") or "error",
                "suggestion": finding.get("suggestion"),
                "source": finding.get("name"),
            }
        )
    for issue in (visual or {}).get("issues") or []:
        if issue.get("severity") == "info" and issue.get("type") == "chapter_midpage":
            continue
        steps.append(visual_step(issue))
    auto_applyable = [
        step
        for step in steps
        if step["action"]
        in {
            "set_default_run",
            "set_style",
            "set_page",
            "set_page_number",
            "set_header_footer",
            "set_keep_next",
            "scale_drawing",
        }
        and step["property"] != "exists"
    ]
    needs_confirmation = [
        step
        for step in steps
        if step["action"] in {"map_structure", "inspect", "set_page_break_before"} or step["property"] == "exists"
    ]
    return {
        "ok": not steps,
        "input": diff.get("input"),
        "profile": diff.get("profile"),
        "compliancePercent": diff.get("compliancePercent"),
        "categoryScores": diff.get("categoryScores"),
        "repairPlan": steps,
        "autoApplyable": len(auto_applyable),
        "needsConfirmation": needs_confirmation,
        "visualIssueCount": len((visual or {}).get("issues") or []),
        "summary": {
            "failed": len(steps),
            "autoApplyable": len(auto_applyable),
            "needsConfirmation": len(needs_confirmation),
            "visualIssues": len((visual or {}).get("issues") or []),
        },
        "notes": [
            "Repair plans are explanations, not silent mutations.",
            "Run apply only after the user confirms the plan.",
            "Visual keepNext/scale_drawing steps are conservative; pageBreakBefore still needs confirmation.",
            "Structural gaps still need paragraphRules or a template, not blind OOXML edits.",
        ],
    }


def render_text(plan: dict[str, Any]) -> str:
    lines = [
        "Repair Plan",
        "────────────────────────",
        f"Failed checks: {plan['summary']['failed']}",
        f"Auto-applyable: {plan['summary']['autoApplyable']}",
        f"Needs confirmation: {plan['summary']['needsConfirmation']}",
        "",
    ]
    for index, step in enumerate(plan.get("repairPlan") or [], start=1):
        lines.append(f"{index}. {step['action']}  {step['target']}.{step['property']}")
        lines.append(f"   from: {step['from']}")
        lines.append(f"   to:   {step['to']}")
        if step.get("affected"):
            lines.append(f"   affected: {step['affected']}")
        if step.get("suggestion"):
            lines.append(f"   {step['suggestion']}")
    if not plan.get("repairPlan"):
        lines.append("No repairs required.")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a repair plan from a DOCX and profile.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--profile", required=True, type=Path)
    parser.add_argument("--output", type=Path, help="Write JSON repair plan")
    parser.add_argument("--text", action="store_true", help="Print a human-readable plan")
    args = parser.parse_args()
    try:
        profile = json.loads(args.profile.read_text(encoding="utf-8"))
        if not isinstance(profile, dict):
            raise ValueError("Profile root must be an object")
        diff = diff_document(args.input, profile)
        visual = visual_qa(args.input, profile, None)
        plan = build_repair_plan(diff, visual)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json.dumps(plan, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        if args.text or not args.output:
            print(render_text(plan), end="")
        else:
            print(json.dumps({"ok": plan["ok"], "output": str(args.output.resolve())}, ensure_ascii=False))
        return 0 if plan["ok"] else 2
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
