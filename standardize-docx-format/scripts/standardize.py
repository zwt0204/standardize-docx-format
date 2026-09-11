#!/usr/bin/env python3
"""Unified CLI for the document formatting compiler pipeline."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from analyze_template import analyze_template
from apply_profile import apply_with_optional_plan, parse_id_list
from audit_docx import audit
from compile_requirements import compile_text, read_source
from diff_profile import diff_document, render_text as render_diff
from document_model import build_model
from profile_schema import validate_profile
from repair_plan import build_repair_plan, render_text as render_plan
from report import write_compliance_report
from validate_docx import validate_docx
from visual_qa import visual_qa, write_html_report
from visual_repair import apply_visual_repairs


def dump(path: Path | None, payload: Any) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    if isinstance(payload, str):
        path.write_text(payload, encoding="utf-8")
    else:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def load_json(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"{path} is not a JSON object")
    return data


def add_io(parser: argparse.ArgumentParser, input_help: str, profile: bool = False, output: bool = True) -> None:
    parser.add_argument("--input", required=True, type=Path, help=input_help)
    if profile:
        parser.add_argument("--profile", required=True, type=Path)
    if output:
        parser.add_argument("--output", type=Path)


def command_audit(args: argparse.Namespace) -> dict[str, Any]:
    report = audit(args.input)
    dump(args.output, report)
    return {"ok": True, "output": str(args.output.resolve()) if args.output else None, "sections": len(report.get("sections") or [])}


def command_model(args: argparse.Namespace) -> dict[str, Any]:
    profile = load_json(args.profile) if getattr(args, "profile", None) else None
    model = build_model(args.input, profile)
    dump(args.output, model)
    return {"ok": True, "summary": model.get("summary"), "output": str(args.output.resolve()) if args.output else None}


def command_diff(args: argparse.Namespace) -> dict[str, Any]:
    profile = load_json(args.profile)
    result = diff_document(args.input, profile)
    dump(args.output, result)
    if args.text:
        print(render_diff(result), end="")
    return result


def command_plan(args: argparse.Namespace) -> dict[str, Any]:
    profile = load_json(args.profile)
    visual = visual_qa(args.input, profile, None)
    plan = build_repair_plan(diff_document(args.input, profile), visual)
    dump(args.output, plan)
    if args.text:
        print(render_plan(plan), end="")
    return plan


def command_apply(args: argparse.Namespace) -> dict[str, Any]:
    return apply_with_optional_plan(
        args.input,
        args.output,
        args.profile,
        args.force,
        args.allow_unresolved,
        plan_path=getattr(args, "plan", None),
        only_ids=parse_id_list(getattr(args, "only_ids", None)),
        skip_ids=parse_id_list(getattr(args, "skip_ids", None)),
        only_auto=bool(getattr(args, "only_auto", False)),
        include_page_breaks=bool(getattr(args, "include_page_breaks", False)),
        confirmed=bool(getattr(args, "yes", False)),
    )


def command_schema(args: argparse.Namespace) -> dict[str, Any]:
    profile = load_json(args.input)
    errors = validate_profile(profile)
    payload = {"ok": not errors, "input": str(args.input.resolve()), "errors": errors}
    dump(args.output, payload)
    return payload


def command_validate(args: argparse.Namespace) -> dict[str, Any]:
    result = validate_docx(args.input, args.profile)
    dump(args.output, result)
    return result


def command_visual(args: argparse.Namespace) -> dict[str, Any]:
    profile = load_json(args.profile) if args.profile else None
    result = visual_qa(args.input, profile, args.render_dir)
    dump(args.output, result)
    if args.html:
        write_html_report(result, args.html)
    return result


def command_compile(args: argparse.Namespace) -> dict[str, Any]:
    result = compile_text(read_source(args.input), args.input.name)
    dump(args.output, result)
    if args.profile_out:
        dump(args.profile_out, result["profile"])
    return result


def command_analyze(args: argparse.Namespace) -> dict[str, Any]:
    result = analyze_template(args.input)
    dump(args.output, result)
    if args.profile_out:
        dump(args.profile_out, result["profile"])
    return result


def command_repair_visual(args: argparse.Namespace) -> dict[str, Any]:
    profile = load_json(args.profile) if args.profile else None
    visual = visual_qa(args.input, profile, None)
    return apply_visual_repairs(args.input, args.output, visual.get("issues") or [], args.force)


def command_pipeline(args: argparse.Namespace) -> dict[str, Any]:
    work = args.work_dir or (args.output.parent if args.output else Path("qa-output"))
    work.mkdir(parents=True, exist_ok=True)
    profile = load_json(args.profile)
    before = audit(args.input)
    dump(work / "before-audit.json", before)
    model = build_model(args.input, profile)
    dump(work / "document-model.json", model)
    diff = diff_document(args.input, profile)
    dump(work / "diff.json", diff)
    visual_before = visual_qa(args.input, profile, work / "rendered-before" if args.render else None)
    dump(work / "visual-before.json", visual_before)
    plan = build_repair_plan(diff, visual_before)
    dump(work / "repair-plan.json", plan)
    applied = None
    visual_applied = None
    validation = None
    visual = visual_before
    output_docx = args.output
    current = args.input
    if args.apply:
        if not output_docx:
            raise ValueError("pipeline --apply requires --output for the standardized DOCX")
        if not args.yes:
            raise ValueError("Refusing to apply without --yes after reviewing the repair plan")
        plan_path = work / "repair-plan.json"
        applied = apply_with_optional_plan(
            args.input,
            output_docx,
            args.profile,
            args.force,
            args.allow_unresolved,
            plan_path=plan_path if args.plan_apply else None,
            only_ids=parse_id_list(args.only_ids),
            skip_ids=parse_id_list(args.skip_ids),
            only_auto=args.only_auto,
            include_page_breaks=args.include_page_breaks,
            confirmed=True,
        )
        dump(work / "apply.json", applied)
        current = Path(applied.get("output") or output_docx)
        output_docx = current
        if args.visual_repair and not applied.get("visualRepair"):
            repaired_path = output_docx.with_name(output_docx.stem + ".visual.docx")
            visual_applied = apply_visual_repairs(output_docx, repaired_path, visual_qa(output_docx, profile, None).get("issues") or [], True)
            dump(work / "visual-repair.json", visual_applied)
            current = repaired_path
            output_docx = repaired_path
        after = audit(current)
        dump(work / "after-audit.json", after)
        validation = validate_docx(current, args.profile)
        dump(work / "validation.json", validation)
        visual = visual_qa(current, profile, work / "rendered" if args.render else None)
        dump(work / "visual.json", visual)
        write_html_report(visual, work / "visual-report.html")
    else:
        dump(work / "visual.json", visual)
        write_html_report(visual, work / "visual-report.html")

    compliance = write_compliance_report(work / "report.html", diff, visual, plan, validation)
    dump(work / "compliance.json", compliance)
    report = {
        "ok": bool(diff.get("ok")) and (validation is None or validation.get("ok")) and (visual is None or visual.get("ok")),
        "input": str(args.input.resolve()),
        "profile": str(args.profile.resolve()),
        "output": str(output_docx.resolve()) if output_docx else None,
        "workDir": str(work.resolve()),
        "compliancePercent": diff.get("compliancePercent"),
        "categoryScores": diff.get("categoryScores"),
        "repairPlan": plan.get("summary"),
        "validation": {"ok": validation.get("ok"), "failed": len(validation.get("failed") or [])} if validation else None,
        "visual": {
            "ok": visual.get("ok"),
            "errorCount": visual.get("errorCount"),
            "warningCount": visual.get("warningCount"),
            "pageCount": (visual.get("pagination") or {}).get("pageCount"),
        }
        if visual
        else None,
        "applied": bool(applied),
        "visualRepaired": bool(visual_applied),
        "mustReview": compliance.get("mustReview"),
        "estimatedVisual": compliance.get("estimated"),
    }
    dump(work / "pipeline-report.json", report)
    return report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="DOCX formatting compiler: audit, diff, repair, apply, visual QA.")
    sub = parser.add_subparsers(dest="command", required=True)

    audit_p = sub.add_parser("audit", help="Read-only package and style inventory")
    add_io(audit_p, "Source .docx")
    audit_p.set_defaults(func=command_audit)

    model_p = sub.add_parser("model", help="Build a semantic Document AST")
    add_io(model_p, "Source .docx")
    model_p.add_argument("--profile", type=Path)
    model_p.set_defaults(func=command_model)

    diff_p = sub.add_parser("diff", help="Explain expected vs actual formatting")
    add_io(diff_p, "Source .docx", profile=True)
    diff_p.add_argument("--text", action="store_true")
    diff_p.set_defaults(func=command_diff)

    plan_p = sub.add_parser("plan", help="Build a confirmable repair plan")
    add_io(plan_p, "Source .docx", profile=True)
    plan_p.add_argument("--text", action="store_true")
    plan_p.set_defaults(func=command_plan)

    apply_p = sub.add_parser("apply", help="Apply a profile to a new DOCX")
    apply_p.add_argument("--input", required=True, type=Path)
    apply_p.add_argument("--output", required=True, type=Path)
    apply_p.add_argument("--profile", required=True, type=Path)
    apply_p.add_argument("--force", action="store_true")
    apply_p.add_argument("--allow-unresolved", action="store_true")
    apply_p.add_argument("--plan", type=Path, help="Confirmed repair-plan JSON")
    apply_p.add_argument("--only-ids", help="Comma-separated repair step ids")
    apply_p.add_argument("--skip-ids", help="Comma-separated repair step ids to skip")
    apply_p.add_argument("--only-auto", action="store_true")
    apply_p.add_argument("--include-page-breaks", action="store_true")
    apply_p.add_argument("--yes", action="store_true", help="Confirm applying a repair plan")
    apply_p.set_defaults(func=command_apply)

    schema_p = sub.add_parser("schema", help="Validate a profile against profile.schema.json")
    schema_p.add_argument("--input", required=True, type=Path)
    schema_p.add_argument("--output", type=Path)
    schema_p.set_defaults(func=command_schema)

    validate_p = sub.add_parser("validate", help="Structural validation against a profile")
    add_io(validate_p, "DOCX to validate", profile=True)
    validate_p.set_defaults(func=command_validate)

    visual_p = sub.add_parser("visual", help="Visual QA heuristics and optional render")
    add_io(visual_p, "Source .docx")
    visual_p.add_argument("--profile", type=Path)
    visual_p.add_argument("--html", type=Path)
    visual_p.add_argument("--render-dir", type=Path)
    visual_p.set_defaults(func=command_visual)

    repair_visual_p = sub.add_parser("repair-visual", help="Apply conservative keepNext / image-scale repairs")
    repair_visual_p.add_argument("--input", required=True, type=Path)
    repair_visual_p.add_argument("--output", required=True, type=Path)
    repair_visual_p.add_argument("--profile", type=Path)
    repair_visual_p.add_argument("--force", action="store_true")
    repair_visual_p.set_defaults(func=command_repair_visual)

    compile_p = sub.add_parser("compile", help="Compile a specification into a draft profile")
    add_io(compile_p, "Specification text/PDF")
    compile_p.add_argument("--profile-out", type=Path)
    compile_p.set_defaults(func=command_compile)

    analyze_p = sub.add_parser("analyze-template", help="Analyze a DOCX template into a draft profile")
    add_io(analyze_p, "Template .docx")
    analyze_p.add_argument("--profile-out", type=Path)
    analyze_p.set_defaults(func=command_analyze)

    pipeline_p = sub.add_parser("pipeline", help="Audit → diff → repair plan → optional apply → visual QA")
    pipeline_p.add_argument("--input", required=True, type=Path)
    pipeline_p.add_argument("--profile", required=True, type=Path)
    pipeline_p.add_argument("--output", type=Path, help="Standardized DOCX path when using --apply")
    pipeline_p.add_argument("--work-dir", type=Path)
    pipeline_p.add_argument("--apply", action="store_true")
    pipeline_p.add_argument("--yes", action="store_true", help="Confirm the repair plan and allow apply")
    pipeline_p.add_argument("--force", action="store_true")
    pipeline_p.add_argument("--allow-unresolved", action="store_true")
    pipeline_p.add_argument("--render", action="store_true")
    pipeline_p.add_argument("--visual-repair", action="store_true", help="After apply, also run conservative visual repairs")
    pipeline_p.add_argument("--plan-apply", action="store_true", help="Apply only selected repair-plan steps instead of the full profile")
    pipeline_p.add_argument("--only-ids", help="Comma-separated repair step ids when using --plan-apply")
    pipeline_p.add_argument("--skip-ids", help="Comma-separated repair step ids to skip")
    pipeline_p.add_argument("--only-auto", action="store_true", help="With --plan-apply, apply only auto-applyable steps")
    pipeline_p.add_argument("--include-page-breaks", action="store_true")
    pipeline_p.set_defaults(func=command_pipeline)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        result = args.func(args)
        if args.command in {"diff", "plan"} and getattr(args, "text", False):
            return 0 if result.get("ok") else 2
        public_keys = (
            "ok",
            "output",
            "input",
            "profile",
            "compliancePercent",
            "errorCount",
            "warningCount",
            "summary",
            "workDir",
            "applied",
            "categoryScores",
            "validation",
            "visual",
            "repairPlan",
            "unresolved",
            "conflicts",
            "appliedCount",
            "visualRepaired",
            "errors",
            "appliedStepIds",
            "appliedStepCount",
        )
        public = {key: result[key] for key in public_keys if key in result}
        if isinstance(public.get("profile"), dict):
            public.pop("profile")
        if args.command == "plan" and isinstance(public.get("repairPlan"), list):
            public["repairPlan"] = result.get("summary")
        public.setdefault("ok", True)
        print(json.dumps(public, ensure_ascii=False, indent=2))
        if args.command in {"diff", "validate", "plan", "visual", "pipeline"}:
            return 0 if result.get("ok") else 2
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
