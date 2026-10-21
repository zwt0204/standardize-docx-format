#!/usr/bin/env python3
"""Validate a DOCX formatting profile against the bundled JSON Schema.

The validator implements the subset of JSON Schema used by profile.schema.json
(type, properties, required, enum, items, additionalProperties, $ref, oneOf,
min/max, minLength, minItems/maxItems). It does not require jsonschema.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

SCHEMA_PATH = Path(__file__).resolve().parents[1] / "references" / "profile.schema.json"


class ProfileSchemaError(ValueError):
    def __init__(self, errors: list[str]) -> None:
        self.errors = errors
        super().__init__("Profile failed schema validation:\n" + "\n".join(f"- {item}" for item in errors))


def load_schema(path: Path | None = None) -> dict[str, Any]:
    schema_path = path or SCHEMA_PATH
    data = json.loads(schema_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"{schema_path} is not a JSON object")
    return data


def _resolve_ref(schema: dict[str, Any], root: dict[str, Any]) -> dict[str, Any]:
    ref = schema.get("$ref")
    if not ref:
        return schema
    if not isinstance(ref, str) or not ref.startswith("#/"):
        raise ValueError(f"Unsupported $ref: {ref}")
    current: Any = root
    for part in ref[2:].split("/"):
        current = current[part]
    if not isinstance(current, dict):
        raise ValueError(f"$ref {ref} did not resolve to an object")
    merged = dict(current)
    for key, value in schema.items():
        if key != "$ref":
            merged[key] = value
    return merged


def _is_type(value: Any, expected: str) -> bool:
    if expected == "object":
        return isinstance(value, dict)
    if expected == "array":
        return isinstance(value, list)
    if expected == "string":
        return isinstance(value, str)
    if expected == "integer":
        return isinstance(value, int) and not isinstance(value, bool)
    if expected == "number":
        return isinstance(value, (int, float)) and not isinstance(value, bool)
    if expected == "boolean":
        return isinstance(value, bool)
    if expected == "null":
        return value is None
    return True


def _validate(instance: Any, schema: dict[str, Any], root: dict[str, Any], path: str, errors: list[str]) -> None:
    schema = _resolve_ref(schema, root)
    if "oneOf" in schema:
        matched = 0
        nested_errors: list[str] = []
        for option in schema["oneOf"]:
            trial: list[str] = []
            _validate(instance, option, root, path, trial)
            if not trial:
                matched += 1
            else:
                nested_errors.extend(trial)
        if matched != 1:
            errors.append(f"{path}: value did not match exactly one oneOf option")
        return

    expected_type = schema.get("type")
    if isinstance(expected_type, str) and not _is_type(instance, expected_type):
        errors.append(f"{path}: expected {expected_type}, got {type(instance).__name__}")
        return
    if isinstance(expected_type, list) and not any(_is_type(instance, item) for item in expected_type):
        errors.append(f"{path}: expected one of {expected_type}, got {type(instance).__name__}")
        return

    if "enum" in schema and instance not in schema["enum"]:
        errors.append(f"{path}: {instance!r} is not one of {schema['enum']}")

    if isinstance(instance, str) and "minLength" in schema and len(instance) < int(schema["minLength"]):
        errors.append(f"{path}: string shorter than minLength {schema['minLength']}")

    if isinstance(instance, (int, float)) and not isinstance(instance, bool):
        if "minimum" in schema and instance < schema["minimum"]:
            errors.append(f"{path}: {instance} is below minimum {schema['minimum']}")
        if "maximum" in schema and instance > schema["maximum"]:
            errors.append(f"{path}: {instance} is above maximum {schema['maximum']}")

    if isinstance(instance, list):
        if "minItems" in schema and len(instance) < int(schema["minItems"]):
            errors.append(f"{path}: array shorter than minItems {schema['minItems']}")
        if "maxItems" in schema and len(instance) > int(schema["maxItems"]):
            errors.append(f"{path}: array longer than maxItems {schema['maxItems']}")
        item_schema = schema.get("items")
        if isinstance(item_schema, dict):
            for index, item in enumerate(instance):
                _validate(item, item_schema, root, f"{path}[{index}]", errors)

    if isinstance(instance, dict):
        properties = schema.get("properties") or {}
        required = schema.get("required") or []
        for key in required:
            if key not in instance:
                errors.append(f"{path}: missing required property {key!r}")
        additional = schema.get("additionalProperties", True)
        for key, value in instance.items():
            child = f"{path}.{key}" if path else key
            if key in properties:
                _validate(value, properties[key], root, child, errors)
            elif additional is False:
                errors.append(f"{path}: additional property {key!r} is not allowed")
            elif isinstance(additional, dict):
                _validate(value, additional, root, child, errors)


def validate_profile(profile: Any, schema: dict[str, Any] | None = None) -> list[str]:
    if not isinstance(profile, dict):
        return ["$: expected object profile"]
    root = schema or load_schema()
    errors: list[str] = []
    _validate(profile, root, root, "$", errors)
    extra = _semantic_checks(profile)
    errors.extend(extra)
    return errors


def _semantic_checks(profile: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    seen_tables: set[int] = set()
    for index, rule in enumerate(profile.get("tableRules") or []):
        if not isinstance(rule, dict):
            continue
        color = rule.get("color", "000000")
        if isinstance(color, str) and not re.fullmatch(r"#?[0-9A-Fa-f]{6}|auto", color):
            errors.append(f"tableRules[{index}].color: expected six hex digits or auto")
        indexes = rule.get("tableIndexes")
        for table_index in indexes if isinstance(indexes, list) else []:
            if not isinstance(table_index, int):
                continue
            if table_index in seen_tables:
                errors.append(f"tableRules[{index}]: table {table_index} is targeted more than once")
            seen_tables.add(table_index)
        for key in ("outerSizePt", "headerSizePt"):
            value = rule.get(key)
            if isinstance(value, (int, float)) and not isinstance(value, bool) and value * 8 != round(value * 8):
                errors.append(f"tableRules[{index}].{key}: must be a multiple of 0.125 pt")
        widths, width = rule.get("columnWidthsIn"), rule.get("widthIn")
        if (isinstance(widths, list) and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in widths)
                and isinstance(width, (int, float)) and abs(width - sum(widths)) > 1 / 1440):
            errors.append(f"tableRules[{index}]: widthIn must equal the sum of columnWidthsIn")
        paragraph = rule.get("paragraph")
        supported_p = {"alignment", "lineSpacing", "lineSpacingExactPt", "spaceBeforePt", "spaceAfterPt", "leftIn", "rightIn", "firstLineIn", "hangingIn", "firstLineChars", "keepNext", "keepLines", "pageBreakBefore", "widowControl", "outlineLevel"}
        run = rule.get("run")
        supported_r = {"latinFont", "eastAsiaFont", "complexScriptFont", "sizePt", "bold", "italic", "color"}
        for label, config, supported in (("paragraph", paragraph, supported_p), ("run", run, supported_r)):
            if isinstance(config, dict):
                for key in config.keys() - supported:
                    errors.append(f"tableRules[{index}].{label}.{key}: unsupported property")
        if isinstance(paragraph, dict):
            if paragraph.get("alignment") not in {None, "left", "center", "right", "justify"}:
                errors.append(f"tableRules[{index}].paragraph.alignment: unsupported alignment")
            for keys in (("lineSpacing", "lineSpacingExactPt"), ("firstLineIn", "firstLineChars"), ("hangingIn", "firstLineIn"), ("hangingIn", "firstLineChars")):
                if all(key in paragraph for key in keys):
                    errors.append(f"tableRules[{index}].paragraph: conflicting {keys[0]} and {keys[1]}")
            for key in ("lineSpacing", "lineSpacingExactPt", "spaceBeforePt", "spaceAfterPt", "leftIn", "rightIn", "firstLineIn", "hangingIn", "firstLineChars"):
                value = paragraph.get(key)
                if isinstance(value, (int, float)) and (value < 0 or (key in {"lineSpacing", "lineSpacingExactPt"} and value == 0)):
                    errors.append(f"tableRules[{index}].paragraph.{key}: invalid negative or zero value")
        if isinstance(run, dict):
            value = run.get("sizePt")
            if isinstance(value, (int, float)) and (value <= 0 or value * 2 != round(value * 2)):
                errors.append(f"tableRules[{index}].run.sizePt: expected a positive multiple of 0.5 pt")
            color = run.get("color")
            if isinstance(color, str) and not re.fullmatch(r"#?[0-9A-Fa-f]{6}", color):
                errors.append(f"tableRules[{index}].run.color: expected six hex digits")
    numbering = profile.get("headingNumbering") or {}
    if isinstance(numbering, dict) and numbering.get("enabled") is True:
        styles = numbering.get("styles")
        patterns = numbering.get("patterns")
        formats = numbering.get("formats")
        starts = numbering.get("starts")
        arrays = [styles, patterns, formats, starts]
        if not all(isinstance(item, list) for item in arrays):
            errors.append("headingNumbering: enabled numbering needs styles, patterns, formats, and starts arrays")
        elif len({len(item) for item in arrays}) != 1:
            errors.append("headingNumbering: styles/patterns/formats/starts must have the same length")
        elif isinstance(patterns, list):
            for index, pattern in enumerate(patterns):
                if isinstance(pattern, str) and f"%{index + 1}" not in pattern:
                    errors.append(f"headingNumbering.patterns[{index}] must include %{index + 1}")
    caption = profile.get("captionNumbering") or {}
    if isinstance(caption, dict) and caption.get("strategy") == "needs-confirmation":
        unresolved = ((profile.get("requirements") or {}).get("unresolved") or [])
        if not unresolved:
            errors.append("captionNumbering.strategy=needs-confirmation requires requirements.unresolved to record the chapter-number policy")
    for index, rule in enumerate(profile.get("paragraphRules") or []):
        if not isinstance(rule, dict):
            continue
        match = rule.get("match") or {}
        if rule.get("textRegexReplace") and isinstance(match, dict):
            style_guard = match.get("styleNameNotRegex") or match.get("currentStyleNotIn") or match.get("styleNameRegex")
            if not style_guard:
                errors.append(
                    f"paragraphRules[{index}]: number-stripping rules should exclude TOC styles via styleNameNotRegex or currentStyleNotIn"
                )
    return errors


def ensure_profile(profile: Any, *, schema: dict[str, Any] | None = None) -> dict[str, Any]:
    errors = validate_profile(profile, schema)
    if errors:
        raise ProfileSchemaError(errors)
    assert isinstance(profile, dict)
    return profile


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a standardization profile against profile.schema.json")
    parser.add_argument("--input", required=True, type=Path, help="Profile JSON")
    parser.add_argument("--schema", type=Path, help="Override schema path")
    args = parser.parse_args()
    try:
        profile = json.loads(args.input.read_text(encoding="utf-8"))
        errors = validate_profile(profile, load_schema(args.schema) if args.schema else None)
        if errors:
            print(json.dumps({"ok": False, "errors": errors}, ensure_ascii=False, indent=2))
            return 2
        print(json.dumps({"ok": True, "input": str(args.input.resolve())}, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
