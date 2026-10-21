"""Layout regressions use OOXML fixtures, not screenshots as proof of acceptance."""
import subprocess
import zipfile
from pathlib import Path
from unittest.mock import patch

import pytest

from document_model import build_model, measure_width
from layout_estimate import paginate
from test_v2_pipeline import W, WP, write_docx
from visual_qa import render_document, visual_qa
from xml.dom import minidom


def source(tmp_path, body):
    path = tmp_path / "layout.docx"
    write_docx(path)
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["word/document.xml"] = f'<w:document xmlns:w="{W}" xmlns:wp="{WP}"><w:body>{body}</w:body></w:document>'.encode()
    with zipfile.ZipFile(path, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return path


def section(width=8, height=10):
    return f'<w:sectPr><w:pgSz w:w="{width * 1440}" w:h="{height * 1440}"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr>'


def paragraph(text="text"):
    return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'


def table(width, kind="dxa", content=None):
    content = content or paragraph()
    return f'<w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="{kind}"/></w:tblPr><w:tr><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr></w:tbl>'


def test_width_units_section_and_nested_container(tmp_path):
    # First section is narrow; a later wide section must not hide its overflow.
    drawing = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="2743200" cy="914400"/></wp:inline></w:drawing></w:r></w:p>'
    nested = table(4000, "pct", drawing)  # 80% of a 2in cell = 1.6in
    path = source(tmp_path, table(5000, "pct", nested) + table(0, "auto") + table(7200) + f'<w:p><w:pPr>{section(6)}</w:pPr></w:p>' + paragraph() + section(12))
    model = build_model(path)
    assert [t["widthIn"] for t in model["tables"]] == pytest.approx([4, 1.6, 4, 5])
    issues = visual_qa(path, None, None)["issues"]
    assert [i["tableIndex"] for i in issues if i["type"] == "table_overflow"] == [3]
    image = next(i for i in issues if i["type"] == "image_overflow")
    assert image["contentWidthIn"] == 2
    assert model["blocks"][0]["inTable"]
    node = minidom.parseString(f'<w:tblW xmlns:w="{W}" w:w="80%" w:type="pct"/>').documentElement
    assert measure_width(node, 5) == 4


def layout_model(rows, content_height=2):
    blocks = []
    table_rows = []
    for row_index, row in enumerate(rows):
        cells = []
        for count in row.get("cells", [10, 10]):
            index = len(blocks)
            blocks.append({"index": index, "type": "table_cell", "section": 0, "textLength": count,
                           "sizePt": 12, "lineSpacing": 1, "topTableIndex": 0})
            cells.append({"widthIn": 2, "paragraphIndexes": [index]})
        table_rows.append({**row, "cells": cells})
    return {"blocks": blocks, "tables": [{"index": 0, "rows": table_rows}],
            "audit": {"sections": [{"page": {"widthIn": 6, "heightIn": content_height + 2,
                       "marginsIn": {"left": 1, "right": 1, "top": 1, "bottom": 1}}}]}}


def test_parallel_cells_use_row_maximum_and_cell_width():
    model = layout_model([{}, {}])
    estimate = paginate(model)
    assert len(estimate["blocks"]) == 2
    assert estimate["pages"][0]["usedHeightIn"] == pytest.approx(2 * (12 / 72 + 0.04), abs=0.001)
    narrow = layout_model([{"cells": [60, 10]}])
    assert paginate(narrow)["blocks"][0]["heightIn"] > estimate["blocks"][0]["heightIn"]


def test_repeated_headers_cant_split_and_tall_rows():
    rows = [{"header": True, "minHeightIn": 0.4}, {"cantSplit": True, "minHeightIn": 1.3},
            {"cantSplit": True, "minHeightIn": 1.3}]
    result = paginate(layout_model(rows))
    assert result["blocks"][-1]["page"] == 2
    assert result["blocks"][-1]["yIn"] == 0.4
    tall = paginate(layout_model([{"cantSplit": True, "minHeightIn": 3}]))
    assert tall["issues"][0]["type"] == "table_row_too_tall"
    split = paginate(layout_model([{"minHeightIn": 3}]))
    assert split["pageCount"] == 2
    assert split["blocks"][0]["splitAcrossPages"]
    assert split["estimated"]


def test_requested_render_unavailable_is_not_success(tmp_path):
    path = source(tmp_path, paragraph() + section())
    with patch("visual_qa.shutil.which", return_value=None):
        result = visual_qa(path, None, tmp_path / "renders")
    assert result["ok"] is False
    assert result["render"]["status"] == "unavailable"
    assert result["heuristicStatus"] != "failed"
    assert result["visualConfirmation"]["status"] == "pending"
    no_render = visual_qa(path, None, None)
    assert no_render["render"]["status"] == "not_requested"


def test_old_pdf_cannot_satisfy_successful_process(tmp_path):
    output = tmp_path / "render"
    output.mkdir()
    (output / "layout.pdf").write_bytes(b"old pdf")
    completed = subprocess.CompletedProcess(["fake"], 0, "", "")
    with patch("visual_qa.shutil.which", side_effect=lambda name: "/fake/soffice" if name == "soffice" else None), patch("visual_qa.run_command", return_value=completed):
        result = render_document(tmp_path / "layout.docx", output)
    assert result["ok"] is False
    assert result["pdf"] is None
    assert Path(result["outputDir"]) != output


def test_render_failure_and_success_are_distinct(tmp_path):
    path = source(tmp_path, paragraph() + section())
    def fake_convert(command, cwd=None):
        target = Path(command[command.index("--outdir") + 1])
        (target / "layout.pdf").write_bytes(b"%PDF-1.7\n" + b"x" * 1500)
        return subprocess.CompletedProcess(command, 0, "", "")
    with patch("visual_qa.shutil.which", side_effect=lambda name: "/fake/soffice" if name == "soffice" else None), patch("visual_qa.run_command", side_effect=fake_convert):
        result = visual_qa(path, None, tmp_path / "renders")
    assert result["ok"]
    assert result["render"]["status"] == "succeeded"
    assert result["visualConfirmation"]["status"] == "pending"
    with patch("visual_qa.render_document", return_value={"ok": False, "status": "failed"}):
        result = visual_qa(path, None, tmp_path / "renders")
    assert not result["ok"]
    assert result["heuristicStatus"] != "failed"


def test_missing_image_location_does_not_resize_unrelated_images(tmp_path):
    from visual_repair import apply_visual_repairs
    drawing = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="2743200" cy="914400"/></wp:inline></w:drawing></w:r></w:p>'
    path = source(tmp_path, drawing + section())
    output = tmp_path / "output.docx"
    result = apply_visual_repairs(path, output, [{"type": "image_overflow", "contentWidthIn": 1}])
    assert result["appliedCount"] == 0
    assert result["skippedCount"] == 1
    assert build_model(output)["blocks"][0]["drawingWidthIn"] == 3


def test_compatibility_branch_indexes_cannot_mutate_wrong_paragraph(tmp_path):
    from visual_repair import apply_visual_repairs
    mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
    branch = f'<mc:AlternateContent xmlns:mc="{mc}"><mc:Choice Requires="w">{paragraph("choice")}</mc:Choice><mc:Fallback>{paragraph("fallback")}</mc:Fallback></mc:AlternateContent>'
    path = source(tmp_path, branch + paragraph("target") + section())
    before = path.read_bytes()
    output = tmp_path / "repaired.docx"
    with pytest.raises(ValueError, match="markup compatibility"):
        apply_visual_repairs(path, output, [{"type": "orphan_heading", "paragraphIndex": 1}])
    assert not output.exists()
    assert path.read_bytes() == before
