#!/usr/bin/env python3
"""Visual QA for DOCX: optional render plus layout heuristics that do not require Word."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any

from audit_docx import attr, direct_child, parse_xml, twips_to_inches
from document_model import build_model
from layout_estimate import detect_pagination_issues, page_content_box, paginate


EMU_PER_INCH = 914400


def add_issue(
    issues: list[dict[str, Any]],
    issue_type: str,
    severity: str,
    message: str,
    **extra: Any,
) -> None:
    payload = {"type": issue_type, "severity": severity, "message": message}
    payload.update(extra)
    issues.append(payload)


def table_widths(document) -> list[float | None]:
    widths: list[float | None] = []
    for tbl in document.getElementsByTagName("w:tbl"):
        tbl_pr = direct_child(tbl, "w:tblPr")
        tbl_w = direct_child(tbl_pr, "w:tblW")
        widths.append(twips_to_inches(attr(tbl_w, "w:w")) if tbl_w is not None else None)
    return widths


def detect_structural_layout_issues(path: Path, model: dict[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    with zipfile.ZipFile(path, "r") as archive:
        document = parse_xml(archive.read("word/document.xml"))
    sections = (model.get("audit") or {}).get("sections") or []
    content_boxes = [page_content_box(section) for section in sections]
    max_content_width = max((box.get("contentWidthIn") or 0) for box in content_boxes) if content_boxes else None

    drawing_index = 0
    for info in model.get("blocks") or []:
        width = info.get("drawingWidthIn")
        if not width:
            continue
        if max_content_width and width > max_content_width + 0.05:
            add_issue(
                issues,
                "image_overflow",
                "error",
                f"图片 {drawing_index + 1} 宽度 {width}in 超出正文区域 {max_content_width}in",
                drawingIndex=drawing_index,
                paragraphIndex=info.get("index"),
                widthIn=width,
                contentWidthIn=max_content_width,
                suggestion="scale drawing to content width",
                repair="scale_drawing",
            )
        drawing_index += 1

    for index, width in enumerate(table_widths(document)):
        if max_content_width and width and width > max_content_width + 0.08:
            add_issue(
                issues,
                "table_overflow",
                "error",
                f"表格 {index + 1} 宽度 {width}in 超出正文区域 {max_content_width}in",
                tableIndex=index,
                widthIn=width,
                contentWidthIn=max_content_width,
                suggestion="narrow the table to the text column",
            )

    blocks = model.get("blocks") or []
    heading_types = {"chapter", "section", "subsection"}
    for offset, info in enumerate(blocks):
        kind = info.get("type")
        following = blocks[offset + 1] if offset + 1 < len(blocks) else None
        previous = blocks[offset - 1] if offset else None
        if kind in heading_types and not info.get("keepNext"):
            add_issue(
                issues,
                "orphan_heading_risk",
                "warning",
                f"标题“{(info.get('preview') or '')[:40]}”未设置 keep_with_next，可能出现页末孤行",
                paragraphIndex=info.get("index"),
                suggestion="keepNext = true",
                repair="set_keep_next",
            )
        if kind == "figure" and following and following.get("type") != "figure_caption":
            add_issue(
                issues,
                "missing_figure_caption",
                "warning",
                f"第 {info.get('index')} 段图片后未紧跟图题",
                paragraphIndex=info.get("index"),
            )
        if kind == "figure_caption" and previous and previous.get("type") not in {"figure", "figure_caption", "body", "empty"}:
            add_issue(
                issues,
                "caption_detached",
                "warning",
                f"图题“{info.get('preview')}”可能与图片分离",
                paragraphIndex=info.get("index"),
            )
        if kind == "table_caption":
            nearby = [previous, following]
            if not any(item and (item.get("inTable") or item.get("type") == "table_cell") for item in nearby):
                add_issue(
                    issues,
                    "caption_detached",
                    "warning",
                    f"表题“{info.get('preview')}”附近未检测到表格",
                    paragraphIndex=info.get("index"),
                )

    empty_run = 0
    empty_start = None
    for info in blocks:
        if info.get("type") == "empty":
            if empty_run == 0:
                empty_start = info.get("index")
            empty_run += 1
            if empty_run >= 8:
                add_issue(
                    issues,
                    "blank_block",
                    "warning",
                    f"连续空段从第 {empty_start} 段开始，可能产生空白页",
                    paragraphIndex=empty_start,
                )
                empty_run = 0
                empty_start = None
        else:
            empty_run = 0
            empty_start = None

    summary = model.get("summary") or {}
    if summary.get("chapterCount", 0) == 0:
        add_issue(issues, "missing_chapters", "warning", "未识别到章节标题，视觉分页建议无法按章检查")
    return issues


def which(names: list[str]) -> str | None:
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


def run_command(command: list[str], cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(command, cwd=cwd, check=False, capture_output=True, text=True)


def render_document(input_path: Path, output_dir: Path) -> dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    docx_cli = shutil.which("docx")
    if docx_cli:
        target = output_dir / "rendered"
        result = run_command([docx_cli, "render", str(input_path), "--out", str(target)])
        return {
            "engine": "docx-cli",
            "ok": result.returncode == 0,
            "command": result.args,
            "stdout": result.stdout[-2000:],
            "stderr": result.stderr[-2000:],
            "outputDir": str(target),
        }

    soffice = which(["soffice", "libreoffice"])
    if soffice:
        result = run_command(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(input_path)]
        )
        pdfs = sorted(output_dir.glob("*.pdf"))
        pdftoppm = shutil.which("pdftoppm")
        pages: list[str] = []
        if pdftoppm and pdfs:
            prefix = output_dir / "page"
            run_command([pdftoppm, "-png", str(pdfs[0]), str(prefix)])
            pages = [str(path) for path in sorted(output_dir.glob("page*.png"))]
        return {
            "engine": "libreoffice",
            "ok": result.returncode == 0 and bool(pdfs),
            "command": result.args,
            "stdout": result.stdout[-2000:],
            "stderr": result.stderr[-2000:],
            "pdf": str(pdfs[0]) if pdfs else None,
            "pages": pages,
            "outputDir": str(output_dir),
        }

    return {
        "engine": None,
        "ok": False,
        "skipped": True,
        "reason": "Neither docx-cli nor LibreOffice is available; structural visual heuristics still ran.",
        "outputDir": str(output_dir),
    }


def inspect_rendered_pages(render: dict[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    pages = [Path(item) for item in render.get("pages") or [] if Path(item).is_file()]
    if len(pages) >= 2:
        sizes = [path.stat().st_size for path in pages]
        median = sorted(sizes)[len(sizes) // 2]
        for index, size in enumerate(sizes, start=1):
            if median and size < max(2048, median * 0.08):
                add_issue(issues, "blank_page", "warning", f"渲染页 {index} 文件过小，可能是空白页", page=index, bytes=size)
    pdf = render.get("pdf")
    if pdf and Path(pdf).is_file() and Path(pdf).stat().st_size < 1024:
        add_issue(issues, "empty_pdf", "error", "渲染得到的 PDF 过小，可能转换失败")
    return issues


def visual_qa(input_path: Path, profile: dict[str, Any] | None, render_dir: Path | None) -> dict[str, Any]:
    model = build_model(input_path, profile)
    issues = detect_structural_layout_issues(input_path, model)
    estimate = paginate(model)
    issues.extend(detect_pagination_issues(estimate))
    render: dict[str, Any] = {"skipped": True, "reason": "render not requested"}
    if render_dir is not None:
        render = render_document(input_path, render_dir)
        issues.extend(inspect_rendered_pages(render))
    # Prefer page-estimated orphan_heading over the weaker keepNext-only risk when both exist.
    seen_heading = {
        (item.get("type"), item.get("paragraphIndex"))
        for item in issues
        if item.get("type") in {"orphan_heading", "heading_body_split"}
    }
    deduped = []
    for item in issues:
        key = (item.get("type"), item.get("paragraphIndex"))
        if item.get("type") == "orphan_heading_risk" and ("orphan_heading", item.get("paragraphIndex")) in seen_heading:
            continue
        if item.get("type") == "orphan_heading_risk" and ("heading_body_split", item.get("paragraphIndex")) in seen_heading:
            continue
        if key in {(d.get("type"), d.get("paragraphIndex")) for d in deduped}:
            continue
        deduped.append(item)
    issues = deduped
    errors = [item for item in issues if item.get("severity") == "error"]
    warnings = [item for item in issues if item.get("severity") == "warning"]
    return {
        "ok": not errors,
        "input": str(input_path.resolve()),
        "issues": issues,
        "errorCount": len(errors),
        "warningCount": len(warnings),
        "pagination": {
            "engine": estimate.get("engine"),
            "pageCount": estimate.get("pageCount"),
            "notes": estimate.get("notes"),
        },
        "render": render,
        "summary": model.get("summary"),
        "notes": [
            "Package validation is not visual validation.",
            "Pagination estimates catch orphan headings before a renderer is available.",
            "When Word/LibreOffice is missing, this command still reports layout risks from OOXML.",
        ],
    }


def write_html_report(result: dict[str, Any], output: Path) -> None:
    rows = []
    for item in result.get("issues") or []:
        rows.append(
            "<tr>"
            f"<td>{item.get('severity')}</td>"
            f"<td>{item.get('type')}</td>"
            f"<td>{item.get('message')}</td>"
            f"<td>{item.get('page', '')} / {item.get('paragraphIndex', '')}</td>"
            "</tr>"
        )
    if not rows:
        rows.append("<tr><td colspan='4'>No visual issues reported.</td></tr>")
    html = f"""<!doctype html>
<html lang="zh">
<head>
  <meta charset="utf-8"/>
  <title>Visual QA</title>
  <style>
    body {{ font-family: sans-serif; margin: 2rem; }}
    .ok {{ color: #0a7; }}
    .bad {{ color: #c33; }}
    table {{ border-collapse: collapse; width: 100%; }}
    td, th {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
  </style>
</head>
<body>
  <h1>Visual QA</h1>
  <p class="{'ok' if result.get('ok') else 'bad'}">errors={result.get('errorCount')} warnings={result.get('warningCount')}</p>
  <p>render: {(result.get('render') or {}).get('engine') or 'none'}; pagination: {(result.get('pagination') or {}).get('pageCount') or '?'} estimated pages</p>
  <table>
    <thead><tr><th>severity</th><th>type</th><th>message</th><th>where</th></tr></thead>
    <tbody>
      {''.join(rows)}
    </tbody>
  </table>
</body>
</html>
"""
    output.write_text(html, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Run visual QA heuristics and optional rendering.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--profile", type=Path)
    parser.add_argument("--output", type=Path, help="Write JSON visual report")
    parser.add_argument("--html", type=Path, help="Write HTML visual report")
    parser.add_argument("--render-dir", type=Path, help="Attempt PDF/PNG render into this directory")
    parser.add_argument("--render", action="store_true", help="Render next to the JSON output or into ./rendered")
    args = parser.parse_args()
    try:
        profile = None
        if args.profile:
            loaded = json.loads(args.profile.read_text(encoding="utf-8"))
            if not isinstance(loaded, dict):
                raise ValueError("Profile root must be an object")
            profile = loaded
        render_dir = args.render_dir
        if args.render and render_dir is None:
            render_dir = (args.output.parent if args.output else Path("rendered"))
        result = visual_qa(args.input, profile, render_dir)
        rendered = json.dumps(result, ensure_ascii=False, indent=2)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(rendered + "\n", encoding="utf-8")
        if args.html:
            args.html.parent.mkdir(parents=True, exist_ok=True)
            write_html_report(result, args.html)
        if args.output:
            print(json.dumps({"ok": result["ok"], "output": str(args.output.resolve()), "errorCount": result["errorCount"], "warningCount": result["warningCount"]}, ensure_ascii=False))
        else:
            print(rendered)
        return 0 if result["ok"] else 2
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
