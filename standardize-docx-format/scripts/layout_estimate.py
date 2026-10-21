#!/usr/bin/env python3
"""Estimate pagination from OOXML so Visual QA can flag orphan headings without Word."""

from __future__ import annotations

from typing import Any


def page_content_box(section: dict[str, Any]) -> dict[str, float | None]:
    page = section.get("page") or {}
    margins = page.get("marginsIn") or {}
    width = page.get("widthIn")
    height = page.get("heightIn")
    left = margins.get("left") or 0
    right = margins.get("right") or 0
    top = margins.get("top") or 0
    bottom = margins.get("bottom") or 0
    return {
        "widthIn": width,
        "heightIn": height,
        "contentWidthIn": (width - left - right) if width else None,
        "contentHeightIn": (height - top - bottom) if height else None,
        "marginsIn": margins,
    }


HEADING_TYPES = {"chapter", "section", "subsection"}


def chars_per_line(content_width_in: float | None, size_pt: float) -> int:
    width = content_width_in or 6.5
    em = max(size_pt, 9) / 72
    return max(8, int(width / (em * 0.92)))


def estimated_height_in(block: dict[str, Any], content_width_in: float | None, default_size: float, default_spacing: float) -> float:
    if block.get("drawingHeightIn"):
        extra = 0.28 if block.get("type") == "figure" else 0.12
        return float(block["drawingHeightIn"]) + extra
    size = float(block.get("sizePt") or default_size or 12)
    spacing = float(block.get("lineSpacing") or default_spacing or 1.5)
    text_len = int(block.get("textLength") or 0)
    if text_len == 0 and not block.get("hasDrawing"):
        return 0.18 * spacing
    lines = max(1, (text_len + chars_per_line(content_width_in, size) - 1) // chars_per_line(content_width_in, size))
    if block.get("type") in HEADING_TYPES:
        lines = max(1, min(lines, 3))
        return round((size / 72) * spacing * lines + 0.18, 4)
    return round((size / 72) * spacing * lines + 0.04, 4)


def paginate(model: dict[str, Any]) -> dict[str, Any]:
    sections = (model.get("audit") or {}).get("sections") or [{}]
    boxes = [page_content_box(section) for section in sections]
    default_run = (model.get("audit") or {}).get("defaultRun") or {}
    default_size = float(default_run.get("sizePt") or 12)
    styles = (model.get("audit") or {}).get("styles") or {}
    normal = (styles.get("Normal") or {}).get("paragraph") or {}
    default_spacing = float(normal.get("lineSpacing") or 1.5)

    pages: list[dict[str, Any]] = []
    page_index = 1
    section_index = 0
    box = boxes[0] if boxes else page_content_box({})
    content_height = float(box.get("contentHeightIn") or 9.5)
    y = 0.0
    current_page = {"page": page_index, "section": section_index, "blocks": [], "usedHeightIn": 0.0}

    def flush() -> None:
        nonlocal current_page, pages, y
        current_page["usedHeightIn"] = round(y, 4)
        current_page["remainingHeightIn"] = round(max(content_height - y, 0), 4)
        pages.append(current_page)

    def new_page(next_section: int) -> None:
        nonlocal page_index, section_index, box, content_height, y, current_page
        flush()
        page_index += 1
        section_index = next_section
        box = boxes[section_index] if section_index < len(boxes) else box
        content_height = float(box.get("contentHeightIn") or content_height)
        y = 0.0
        current_page = {"page": page_index, "section": section_index, "blocks": [], "usedHeightIn": 0.0}

    source_blocks = model.get("blocks") or []
    by_index = {block["index"]: block for block in source_blocks}
    tables = {table["index"]: table for table in model.get("tables") or []}

    def row_height(row):
        heights = []
        for cell in row.get("cells") or []:
            height = sum(estimated_height_in(by_index[index], cell.get("widthIn"), default_size, default_spacing)
                         for index in cell.get("paragraphIndexes") or [] if index in by_index)
            height += sum(sum(row_height(r) for r in tables[index].get("rows") or [])
                          for index in cell.get("tableIndexes") or [] if index in tables)
            heights.append(height)
        return max([float(row.get("minHeightIn") or 0), *heights, 0.08])

    # Adjacent cells occupy the same vertical band; paragraphs within each cell stack.
    flow = []
    seen_tables = set()
    for block in source_blocks:
        table_index = block.get("topTableIndex")
        if table_index is None or table_index not in tables:
            flow.append(block)
            continue
        if table_index in seen_tables:
            continue
        seen_tables.add(table_index)
        table = tables[table_index]
        header_height = 0.0
        in_header = True
        for row_index, row in enumerate(table.get("rows") or []):
            height = row_height(row)
            if in_header and row.get("header"):
                header_height += height
            else:
                in_header = False
            first_index = next((i for cell in row.get("cells") or [] for i in cell.get("paragraphIndexes") or []), block["index"])
            flow.append({**block, "index": first_index, "type": "table_cell", "keepNext": False,
                         "pageBreakBefore": False, "tableIndex": table_index, "rowIndex": row_index,
                         "rowHeightIn": height, "cantSplit": row.get("cantSplit"),
                         "repeatHeaderHeightIn": 0 if row.get("header") else header_height})

    placed: list[dict[str, Any]] = []
    pagination_issues = []
    for offset, block in enumerate(flow):
        block_section = int(block.get("section") or 0)
        if block_section != section_index:
            new_page(block_section)
        if block.get("pageBreakBefore") and y > 0:
            new_page(block_section)
        height = block.get("rowHeightIn") or estimated_height_in(block, box.get("contentWidthIn"), default_size, default_spacing)
        next_item = flow[offset + 1] if offset + 1 < len(flow) else None
        lookahead = (next_item.get("rowHeightIn") or estimated_height_in(next_item, box.get("contentWidthIn"), default_size, default_spacing)) if next_item else 0
        needed = height + (lookahead if block.get("keepNext") else 0)
        is_row = "rowHeightIn" in block
        header = float(block.get("repeatHeaderHeightIn") or 0)
        if y > 0 and y + needed > content_height + 0.02 and (not is_row or block.get("cantSplit")):
            new_page(block_section)
            if is_row and header < content_height:
                y = header
        if is_row and height + header > content_height + 0.02:
            pagination_issues.append({"type": "table_row_too_tall", "severity": "warning",
                "message": "估算表格行连同重复表头超过一页正文高度，需真实渲染复核",
                "tableIndex": block.get("tableIndex"), "rowIndex": block.get("rowIndex"),
                "paragraphIndex": block.get("index"), "page": page_index})
        remaining_before = content_height - y
        record = {**{key: block.get(key) for key in ("index", "type", "preview", "styleId", "keepNext", "tableIndex", "rowIndex")},
                  "page": page_index, "yIn": round(y, 4), "heightIn": height,
                  "remainingBeforeIn": round(remaining_before, 4)}
        current_page["blocks"].append(record["index"])
        placed.append(record)
        if is_row and not block.get("cantSplit") and y + height > content_height + 0.02:
            record["splitAcrossPages"] = True
            remaining = height
            while remaining > content_height - y + 0.02:
                remaining -= max(content_height - y, 0)
                y = content_height
                new_page(block_section)
                # An oversized header cannot leave space for a split body row.
                y = header if header < content_height - 0.02 else 0
                current_page["blocks"].append(record["index"])
            y += remaining
        else:
            y += height

    flush()
    return {
        "engine": "ooxml-estimate",
        "estimated": True,
        "issues": pagination_issues,
        "pageCount": len(pages),
        "pages": pages,
        "blocks": placed,
        "notes": [
            "This is a deterministic pagination estimate, not Microsoft Word layout.",
            "Use it to catch orphan headings and caption splits before a renderer is available.",
        ],
    }


def detect_pagination_issues(estimate: dict[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = list(estimate.get("issues") or [])
    blocks = estimate.get("blocks") or []
    pages = {page["page"]: page for page in estimate.get("pages") or []}

    for offset, item in enumerate(blocks):
        following = blocks[offset + 1] if offset + 1 < len(blocks) else None
        kind = item.get("type")
        remaining = item.get("remainingBeforeIn") or 0
        heading_height = item.get("heightIn") or 0
        following_height = following.get("heightIn") if following else 0.7
        leftover = remaining - heading_height
        if kind in HEADING_TYPES and leftover < max(following_height, 0.7) - 0.05:
            issues.append(
                {
                    "type": "orphan_heading",
                    "severity": "warning",
                    "message": f"第 {item['page']} 页标题“{(item.get('preview') or '')[:40]}”靠近页底，正文可能落到下一页",
                    "page": item.get("page"),
                    "paragraphIndex": item.get("index"),
                    "suggestion": "keepNext = true",
                    "repair": "set_keep_next",
                }
            )
        if kind in HEADING_TYPES and following and following.get("type") in {"body", "abstract_body"}:
            if following.get("page") != item.get("page"):
                issues.append(
                    {
                        "type": "heading_body_split",
                        "severity": "warning",
                        "message": f"标题“{(item.get('preview') or '')[:40]}”在第 {item['page']} 页，正文在第 {following['page']} 页",
                        "page": item.get("page"),
                        "paragraphIndex": item.get("index"),
                        "suggestion": "keepNext = true",
                        "repair": "set_keep_next",
                    }
                )
        if kind == "figure" and following and following.get("type") == "figure_caption" and following.get("page") != item.get("page"):
            issues.append(
                {
                    "type": "figure_caption_split",
                    "severity": "warning",
                    "message": f"图片在第 {item['page']} 页，图题在第 {following['page']} 页",
                    "page": item.get("page"),
                    "paragraphIndex": item.get("index"),
                    "suggestion": "keepNext = true on the figure paragraph",
                    "repair": "set_keep_next",
                }
            )
        if kind == "table_caption" and following and following.get("type") == "table_cell" and following.get("page") != item.get("page"):
            issues.append(
                {
                    "type": "table_caption_split",
                    "severity": "warning",
                    "message": f"表题在第 {item['page']} 页，表格在第 {following['page']} 页",
                    "page": item.get("page"),
                    "paragraphIndex": item.get("index"),
                    "suggestion": "keepNext = true on the caption",
                    "repair": "set_keep_next",
                }
            )
        if kind == "chapter" and item.get("yIn", 0) > 1.8 and item.get("page", 1) > 1:
            issues.append(
                {
                    "type": "chapter_midpage",
                    "severity": "info",
                    "message": f"章标题“{(item.get('preview') or '')[:40]}”没有另起页",
                    "page": item.get("page"),
                    "paragraphIndex": item.get("index"),
                    "suggestion": "pageBreakBefore = true",
                    "repair": "set_page_break_before",
                }
            )

    for page in pages.values():
        if page.get("usedHeightIn", 0) < 0.4 and page.get("page", 1) > 1:
            issues.append(
                {
                    "type": "blank_page",
                    "severity": "warning",
                    "message": f"估算第 {page['page']} 页几乎空白",
                    "page": page.get("page"),
                    "repair": None,
                }
            )
    return issues
