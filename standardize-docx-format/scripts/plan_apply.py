#!/usr/bin/env python3
"""Turn a confirmed repair plan into a filtered profile plus visual issues."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

AUTO_ACTIONS = {
    "set_default_run",
    "set_style",
    "set_paragraph",
    "set_run",
    "set_page",
    "set_page_number",
    "set_header_footer",
    "set_keep_next",
    "scale_drawing",
}
CONFIRM_ACTIONS = {"map_structure", "inspect", "set_page_break_before"}
VISUAL_ACTIONS = {"set_keep_next", "scale_drawing", "set_page_break_before"}


def step_id(index: int, step: dict[str, Any]) -> str:
    if step.get("id"):
        return str(step["id"])
    source = str(step.get("source") or step.get("property") or index)
    return f"s{index:03d}-{source}"


def is_auto_applyable(step: dict[str, Any]) -> bool:
    action = str(step.get("action") or "")
    return action in AUTO_ACTIONS and step.get("property") != "exists"


def needs_confirmation(step: dict[str, Any]) -> bool:
    action = str(step.get("action") or "")
    return action in CONFIRM_ACTIONS or step.get("property") == "exists"


def annotate_steps(steps: list[dict[str, Any]]) -> list[dict[str, Any]]:
    annotated: list[dict[str, Any]] = []
    for index, step in enumerate(steps, start=1):
        item = dict(step)
        item["id"] = step_id(index, item)
        item["autoApplyable"] = is_auto_applyable(item)
        item["needsConfirmation"] = needs_confirmation(item)
        item.setdefault("reversible", item["action"] not in {"map_structure", "inspect"})
        annotated.append(item)
    return annotated


def select_steps(
    plan: dict[str, Any],
    *,
    only_ids: list[str] | None = None,
    skip_ids: list[str] | None = None,
    only_auto: bool = False,
    include_page_breaks: bool = False,
) -> list[dict[str, Any]]:
    selected: list[dict[str, Any]] = []
    skip = set(skip_ids or [])
    only = set(only_ids or [])
    for step in plan.get("repairPlan") or []:
        identifier = str(step.get("id") or "")
        if identifier in skip:
            continue
        if only and identifier not in only:
            continue
        if only_auto and not step.get("autoApplyable"):
            continue
        if step.get("action") == "set_page_break_before" and not include_page_breaks and identifier not in only:
            continue
        if step.get("action") in CONFIRM_ACTIONS and identifier not in only:
            continue
        selected.append(step)
    return selected


def _keep_style_fields(config: dict[str, Any], wanted: set[str]) -> dict[str, Any]:
    kept: dict[str, Any] = {}
    for bucket in ("run", "paragraph"):
        source = config.get(bucket)
        if not isinstance(source, dict):
            continue
        filtered = {key: value for key, value in source.items() if key in wanted}
        if filtered:
            kept[bucket] = filtered
    for key in ("name", "basedOn", "next"):
        if key in config and key in wanted:
            kept[key] = config[key]
    return kept


def filter_profile_for_steps(profile: dict[str, Any], steps: list[dict[str, Any]]) -> dict[str, Any]:
    """Keep only profile sections corresponding to selected formatting steps.

    Visual-only plans return a formatting-empty profile so apply can still
    copy the package and then visual repair can run.
    """
    filtered = deepcopy(profile)
    formatting_actions = {step.get("action") for step in steps if not step.get("visual")}
    if not formatting_actions:
        for key in (
            "page",
            "sections",
            "defaultRun",
            "styles",
            "headingNumbering",
            "pageNumbers",
            "paragraphRules",
            "captionRules",
            "bookmarkRules",
            "headersFooters",
            "replacements",
            "fieldRules",
            "fields",
        ):
            filtered.pop(key, None)
        return filtered

    if {"set_style", "set_paragraph", "set_run"} & formatting_actions:
        formatting_actions.add("set_style")
    keep_page = "set_page" in formatting_actions
    keep_default = "set_default_run" in formatting_actions
    keep_styles: dict[str, set[str]] = {}
    keep_page_numbers = False
    keep_headers = False
    header_indexes: set[int] = set()
    page_number_sections: set[int] = set()

    for step in steps:
        if step.get("visual"):
            continue
        action = step.get("action")
        source = str(step.get("source") or "")
        target = str(step.get("target") or "")
        prop = str(step.get("property") or "")
        if action in {"set_style", "set_paragraph", "set_run"}:
            style_id = target
            if source.startswith("styles."):
                style_id = source.split(".")[1]
            keep_styles.setdefault(style_id, set()).add(prop)
        elif action == "set_page_number":
            keep_page_numbers = True
            if "section[" in source:
                try:
                    page_number_sections.add(int(source.split("section[", 1)[1].split("]", 1)[0]))
                except ValueError:
                    pass
        elif action == "set_header_footer":
            keep_headers = True
            if source.startswith("headersFooters["):
                try:
                    header_indexes.add(int(source.split("[", 1)[1].split("]", 1)[0]))
                except ValueError:
                    pass

    if not keep_page:
        filtered.pop("page", None)
        filtered.pop("sections", None)
    if not keep_default:
        filtered.pop("defaultRun", None)
    if keep_styles:
        original_styles = profile.get("styles") or {}
        styles_out: dict[str, Any] = {}
        for style_id, wanted in keep_styles.items():
            config = original_styles.get(style_id)
            if not isinstance(config, dict):
                continue
            kept = _keep_style_fields(config, wanted)
            if "exists" in wanted or not kept:
                styles_out[style_id] = deepcopy(config)
            else:
                styles_out[style_id] = kept
        filtered["styles"] = styles_out
    else:
        filtered.pop("styles", None)
    if keep_page_numbers:
        numbers = profile.get("pageNumbers") or []
        if page_number_sections:
            filtered["pageNumbers"] = [
                item for item in numbers if isinstance(item, dict) and item.get("section") in page_number_sections
            ]
    else:
        filtered.pop("pageNumbers", None)
    if keep_headers:
        headers = profile.get("headersFooters") or []
        if header_indexes:
            filtered["headersFooters"] = [headers[index] for index in sorted(header_indexes) if 0 <= index < len(headers)]
    else:
        filtered.pop("headersFooters", None)

    # Paragraph/caption/field mapping is never inferred from a style diff.
    # Keep those arrays only when the selected plan explicitly asked for them.
    keep_mapping = any(
        str(step.get("source") or "").startswith(prefix)
        for step in steps
        for prefix in ("paragraphRules", "captionRules", "bookmarkRules", "replacements", "fieldRules", "headingNumbering")
    )
    if not keep_mapping:
        for key in ("paragraphRules", "captionRules", "bookmarkRules", "replacements", "fieldRules", "headingNumbering"):
            filtered.pop(key, None)
    return filtered


def visual_issues_from_steps(steps: list[dict[str, Any]]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    for step in steps:
        if not step.get("visual") and step.get("action") not in VISUAL_ACTIONS:
            continue
        issue_type = str(step.get("source") or step.get("property") or "visual")
        issue = {
            "type": issue_type,
            "severity": step.get("severity") or "warning",
            "message": step.get("suggestion"),
            "paragraphIndex": step.get("paragraphIndex"),
            "page": step.get("page"),
            "repair": step.get("action"),
            "contentWidthIn": step.get("contentWidthIn"),
            "widthIn": step.get("widthIn"),
            "suggestion": step.get("to") if isinstance(step.get("to"), str) else step.get("suggestion"),
        }
        issues.append(issue)
    return issues
