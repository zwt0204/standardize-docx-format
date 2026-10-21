"""Compare the actual contents of header/footer parts with profile expectations."""
from __future__ import annotations

from typing import Any

RUN_KEYS = {"latinFont": "latin", "eastAsiaFont": "eastAsia", "complexScriptFont": "complexScript",
            "sizePt": "sizePt", "bold": "bold", "italic": "italic", "color": "color"}


def matches(actual: Any, expected: Any) -> bool:
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return isinstance(actual, (int, float)) and abs(actual - expected) <= 0.01
    return actual == expected


def tokens(segments: list[dict[str, Any]]) -> list[tuple[str, str]]:
    """Ignore text run splitting and field caches while retaining content order."""
    out: list[tuple[str, str]] = []
    for segment in segments:
        key = "text" if "text" in segment else "field"
        value = str(segment.get(key, ""))
        if key == "field":
            value = " ".join(value.split())
        if not value:
            continue
        if key == "text" and out and out[-1][0] == "text":
            out[-1] = ("text", out[-1][1] + value)
        else:
            out.append((key, value))
    return out


def marginal_expectation(report: dict[str, Any], rule: dict[str, Any]) -> tuple[bool, dict[str, Any]]:
    sections = report.get("sections") or []
    index = rule.get("section")
    kind = rule.get("kind")
    reference_type = rule.get("type", "default")
    action = rule.get("action", "set")
    if action == "inherit":
        return True, {"preserved": True}
    if not isinstance(index, int) or not 0 <= index < len(sections):
        return False, {"reason": "section does not exist"}
    section = sections[index]
    refs = [ref for ref in section.get(f"{kind}References", []) if ref.get("type") == reference_type]
    inventory = {part["part"]: part for part in (report.get("marginals") or {}).get("parts", [])}
    part = inventory.get(refs[0].get("target")) if len(refs) == 1 else None
    errors: list[str] = []
    if part is None:
        errors.append("missing, ambiguous, or unresolved reference")
    elif part.get("kind") != kind or refs[0].get("relationshipType", "").rsplit("/", 1)[-1] != kind:
        errors.append("reference does not resolve to the expected header/footer kind")
    if reference_type == "first" and not section.get("differentFirstPage"):
        errors.append("different first page is disabled")
    if reference_type == "even" and not (report.get("marginals") or {}).get("evenAndOddHeaders"):
        errors.append("even and odd headers are disabled")
    if part is None:
        return False, {"references": refs, "errors": errors}
    errors.extend("unsupported: " + message for message in part.get("unsupportedLayouts", []))
    actual_paragraphs = part.get("paragraphDetails", [])
    if action == "clear":
        if part.get("hasObjects") or any(tokens(p["segments"]) for p in actual_paragraphs):
            errors.append("header/footer is not empty")
    else:
        if part.get("hasObjects"):
            errors.append("header/footer contains additional visible content not described by the profile")
        expected_paragraphs = rule.get("paragraphs") or [{
            "text": rule.get("text", ""), "segments": rule.get("segments"),
            "alignment": rule.get("alignment", (rule.get("paragraph") or {}).get("alignment", "center")), "paragraph": rule.get("paragraph", {}),
            "run": rule.get("run", {}), "bottomBorder": rule.get("bottomBorder", {})}]
        if len(actual_paragraphs) != len(expected_paragraphs):
            errors.append("paragraph count differs")
        for i, (actual, expected) in enumerate(zip(actual_paragraphs, expected_paragraphs)):
            prefix = f"paragraph[{i}]"
            segments = expected.get("segments")
            if segments is None:
                segments = [{"text": str(expected.get("text", ""))}]
            if tokens(actual["segments"]) != tokens(segments):
                errors.append(f"{prefix} text or fields differ")
            props = dict(expected.get("paragraph") or {})
            if "alignment" in expected:
                props["alignment"] = expected["alignment"]
            for key, value in props.items():
                if key in {"firstLineIn", "firstLineChars", "hangingIn"} and value == 0:
                    # A zero result alone can hide competing inherited units.
                    actual_indent = actual["paragraph"]
                    if any(actual_indent.get(name) not in (None, 0)
                           for name in ("firstLineChars", "hangingChars", "hangingIn")):
                        errors.append(f"{prefix}.{key} has competing nonzero indentation")
                key_actual = key
                if key == "alignment" and value == "justify":
                    value = "both"
                if key == "firstLineChars":
                    key_actual = "firstLineIn"
                    value = value * ((expected.get("run") or {}).get("sizePt") or 12) / 72
                if not matches(actual["paragraph"].get(key_actual), value):
                    errors.append(f"{prefix}.{key} differs")
            border = expected.get("bottomBorder") or {}
            if border:
                border = {"style": "single", "size": 6, "space": 1, "color": "000000", **border}
                for key, value in border.items():
                    value = str(value).lstrip("#").upper() if key == "color" else str(value)
                    if actual["bottomBorder"].get(key) != value:
                        errors.append(f"{prefix}.bottomBorder.{key} differs")
            # Match formatting by text offsets so harmless Word run splitting is accepted.
            expected_units: list[tuple[str, str, dict[str, Any]]] = []
            for segment in segments:
                config = {**(expected.get("run") or {}), **(segment.get("run") or {})}
                if "text" in segment:
                    expected_units.extend(("text", char, config) for char in str(segment["text"]))
                else:
                    expected_units.append(("field", " ".join(str(segment.get("field", "")).split()), config))
            actual_units: list[tuple[str, str, list[dict[str, Any]]]] = []
            for segment in actual["segments"]:
                if "text" in segment:
                    actual_units.extend(("text", char, segment["runs"]) for char in segment["text"])
                else:
                    actual_units.append(("field", segment["field"], segment["runs"]))
            if [(k, v) for k, v, _ in expected_units] == [(k, v) for k, v, _ in actual_units]:
                for (_, _, config), (_, _, runs) in zip(expected_units, actual_units):
                    for key, mapped in RUN_KEYS.items():
                        if key not in config:
                            continue
                        value = config[key]
                        if key == "color":
                            value = str(value).lstrip("#").upper()
                        if not runs or any(not matches(run.get(mapped), value) for run in runs):
                            error = f"{prefix}.run.{key} differs"
                            if error not in errors:
                                errors.append(error)
    return not errors, {"references": refs, "part": part["part"], "errors": errors}
