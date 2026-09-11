#!/usr/bin/env python3
"""Turn a profile diff into an explainable, confirmable repair plan."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from diff_profile import diff_document
from plan_apply import annotate_steps, is_auto_applyable, needs_confirmation
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
    if name.startswith("styles.") or (category in {"paragraph", "run"} and "styles." in name):
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
        "contentWidthIn": issue.get("contentWidthIn"),
        "widthIn": issue.get("widthIn"),
        "visual": True,
        "reversible": action != "inspect",
    }


def build_repair_plan(diff: dict[str, Any], visual: dict[str, Any] | None = None) -> dict[str, Any]:
    steps: list[dict[str, Any]] = []
    for finding in diff.get("failed") or []:
        action = action_for(finding)
        steps.append(
            {
                "action": action,
                "target": target_for(finding),
                "property": property_for(finding),
                "from": finding.get("actual"),
                "to": finding.get("expected"),
                "affected": finding.get("affected"),
                "severity": finding.get("severity") or "error",
                "suggestion": finding.get("suggestion"),
                "source": finding.get("name"),
                "visual": False,
                "reversible": action not in {"map_structure", "inspect"},
            }
        )
    for issue in (visual or {}).get("issues") or []:
        if issue.get("severity") == "info" and issue.get("type") == "chapter_midpage":
            continue
        steps.append(visual_step(issue))
    steps = annotate_steps(steps)
    auto_steps = [step for step in steps if is_auto_applyable(step)]
    confirm_steps = [step for step in steps if needs_confirmation(step)]
    return {
        "ok": not steps,
        "input": diff.get("input"),
        "profile": diff.get("profile"),
        "compliancePercent": diff.get("compliancePercent"),
        "categoryScores": diff.get("categoryScores"),
        "repairPlan": steps,
        "autoApplyable": len(auto_steps),
        "needsConfirmation": confirm_steps,
        "visualIssueCount": len((visual or {}).get("issues") or []),
        "summary": {
            "failed": len(steps),
            "autoApplyable": len(auto_steps),
            "needsConfirmation": len(confirm_steps),
            "visualIssues": len((visual or {}).get("issues") or []),
        },
        "notes": [
            "Each step has a stable id. Apply with --plan and --yes, optionally --only-ids / --only-auto.",
            "map_structure and pageBreakBefore never run unless their ids are listed in --only-ids.",
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
        identifier = step.get("id") or index
        flags = []
        if step.get("autoApplyable"):
            flags.append("auto")
        if step.get("needsConfirmation"):
            flags.append("confirm")
        if step.get("visual"):
            flags.append("visual")
        suffix = f" [{' '.join(flags)}]" if flags else ""
        lines.append(f"{index}. [{identifier}] {step['action']}  {step['target']}.{step['property']}{suffix}")
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
