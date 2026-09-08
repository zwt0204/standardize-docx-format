#!/usr/bin/env python3
"""Render a combined formatting + visual compliance report."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def bar(percent: float, width: int = 20) -> str:
    filled = max(0, min(width, round(percent / 100 * width)))
    return "█" * filled + "░" * (width - filled)


def html_escape(value: Any) -> str:
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def write_compliance_report(
    output: Path,
    diff: dict[str, Any] | None = None,
    visual: dict[str, Any] | None = None,
    plan: dict[str, Any] | None = None,
    validation: dict[str, Any] | None = None,
) -> None:
    scores = dict((diff or {}).get("categoryScores") or {})
    visual_errors = (visual or {}).get("errorCount") or 0
    visual_warnings = (visual or {}).get("warningCount") or 0
    visual_score = 100.0
    if visual_errors or visual_warnings:
        visual_score = max(0.0, 100.0 - visual_errors * 12 - visual_warnings * 3)
    scores.setdefault("visual", round(visual_score, 1))
    overall = (diff or {}).get("compliancePercent")
    if overall is None:
        overall = round(sum(scores.values()) / max(len(scores), 1), 1) if scores else 100.0
    else:
        overall = round((float(overall) * max(len(scores) - 1, 1) + visual_score) / max(len(scores), 1), 1)

    score_rows = "".join(
        f"<tr><td>{html_escape(name)}</td><td class='bar'>{bar(value)}</td><td>{value}%</td></tr>"
        for name, value in scores.items()
    )
    issue_rows = []
    for item in (visual or {}).get("issues") or []:
        issue_rows.append(
            "<tr>"
            f"<td>{html_escape(item.get('severity'))}</td>"
            f"<td>{html_escape(item.get('type'))}</td>"
            f"<td>{html_escape(item.get('message'))}</td>"
            f"<td>{html_escape(item.get('page', ''))} / {html_escape(item.get('paragraphIndex', ''))}</td>"
            "</tr>"
        )
    if not issue_rows:
        issue_rows.append("<tr><td colspan='4'>No visual issues reported.</td></tr>")
    failed = (diff or {}).get("failed") or []
    failed_rows = []
    for item in failed[:40]:
        failed_rows.append(
            "<tr>"
            f"<td>{html_escape(item.get('category'))}</td>"
            f"<td>{html_escape(item.get('name'))}</td>"
            f"<td>{html_escape(item.get('expected'))}</td>"
            f"<td>{html_escape(item.get('actual'))}</td>"
            "</tr>"
        )
    if not failed_rows:
        failed_rows.append("<tr><td colspan='4'>No formatting diffs.</td></tr>")
    plan_summary = (plan or {}).get("summary") or {}
    validation_failed = len((validation or {}).get("failed") or [])
    html = f"""<!doctype html>
<html lang="zh">
<head>
  <meta charset="utf-8"/>
  <title>Formatting Compliance Report</title>
  <style>
    body {{ font-family: sans-serif; margin: 2rem; color: #222; }}
    h1, h2 {{ margin-bottom: .4rem; }}
    .score {{ font-size: 2rem; }}
    .ok {{ color: #0a7; }}
    .bad {{ color: #c33; }}
    table {{ border-collapse: collapse; width: 100%; margin: 1rem 0; }}
    td, th {{ border: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; }}
    td.bar {{ font-family: monospace; letter-spacing: 1px; }}
  </style>
</head>
<body>
  <h1>AI Document Formatting Report</h1>
  <p class="score {'ok' if overall >= 90 else 'bad'}">格式合规率：{overall}%</p>
  <p>repair plan: {html_escape(plan_summary.get('failed', 0))} steps,
     auto-applyable {html_escape(plan_summary.get('autoApplyable', 0))},
     validation failed {html_escape(validation_failed)}</p>
  <h2>Category scores</h2>
  <table>
    <thead><tr><th>category</th><th>bar</th><th>score</th></tr></thead>
    <tbody>{score_rows}</tbody>
  </table>
  <h2>Formatting diffs</h2>
  <table>
    <thead><tr><th>category</th><th>name</th><th>expected</th><th>actual</th></tr></thead>
    <tbody>{''.join(failed_rows)}</tbody>
  </table>
  <h2>Visual issues</h2>
  <p>errors={visual_errors} warnings={visual_warnings} estimated pages={(visual or {}).get('pagination', {}).get('pageCount', '')}</p>
  <table>
    <thead><tr><th>severity</th><th>type</th><th>message</th><th>where</th></tr></thead>
    <tbody>{''.join(issue_rows)}</tbody>
  </table>
</body>
</html>
"""
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(html, encoding="utf-8")
