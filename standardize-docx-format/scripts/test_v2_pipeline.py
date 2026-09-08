#!/usr/bin/env python3
"""Regression checks for the v2 compiler pipeline."""

from __future__ import annotations

import zipfile
from pathlib import Path

from compile_requirements import compile_text
from diff_profile import diff_document
from document_model import build_model
from layout_estimate import detect_pagination_issues, paginate
from repair_plan import build_repair_plan
from visual_qa import visual_qa
from visual_repair import apply_visual_repairs


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def write_docx(path: Path) -> None:
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:wp="{WP}">
  <w:body>
    <w:p><w:r><w:t>本科毕业论文</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>摘要</w:t></w:r></w:p>
    <w:p><w:r><w:t>这是中文摘要正文。</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Abstract</w:t></w:r></w:p>
    <w:p><w:r><w:t>This is the English abstract.</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>目录</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>1 绪论 1</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p>
    <w:p><w:r><w:t>本章主要讨论研究背景。</w:t></w:r></w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="12192000" cy="4572000"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p><w:r><w:t>图 1-1 系统架构图</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p>
    <w:p><w:r><w:t>[1] 示例文献.</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"/>
      <w:pgNumType w:fmt="decimal" w:start="1"/>
    </w:sectPr>
  </w:body>
</w:document>
"""
    styles = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="仿宋" w:cs="Calibri"/>
        <w:sz w:val="21"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="仿宋"/>
      <w:sz w:val="21"/>
    </w:rPr>
    <w:pPr>
      <w:spacing w:line="240" w:lineRule="auto"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/>
      <w:sz w:val="32"/>
      <w:b/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOC1">
    <w:name w:val="toc 1"/>
  </w:style>
</w:styles>
"""
    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""
    rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL}">
  <Relationship Id="rId1" Type="{OFFICE_REL}/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL}">
  <Relationship Id="rId1" Type="{OFFICE_REL}/styles" Target="styles.xml"/>
</Relationships>
"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", styles)
        archive.writestr("word/_rels/document.xml.rels", document_rels)


def test_compile_requirements() -> None:
    spec = """
    纸型为 A4，纵向。
    上边距 2.54cm，下边距 2.54cm，左边距 3.17cm，右边距 3.17cm。
    正文宋体小四，1.5 倍行距，两端对齐，首行缩进 2 字符。
    一级标题黑体三号。
    二级标题黑体四号。
    封面填写线长度适当。
    """
    result = compile_text(spec, "fixture.txt")
    profile = result["profile"]
    assert profile["page"]["size"] == "a4"
    assert set(profile["page"]["marginsIn"]) >= {"top", "bottom", "left", "right"}
    assert profile["defaultRun"]["eastAsiaFont"] == "宋体"
    assert profile["defaultRun"]["sizePt"] == 12
    assert profile["styles"]["Heading1"]["run"]["sizePt"] == 16
    assert profile["styles"]["Normal"]["paragraph"]["lineSpacing"] == 1.5
    assert any("不可量化" in item for item in result["unresolved"])
    assert profile["document"]["type"] == "thesis"


def test_model_diff_visual(tmp_path: Path) -> None:
    docx = tmp_path / "sample.docx"
    write_docx(docx)
    profile = {
        "name": "fixture",
        "page": {"size": "a4", "orientation": "portrait"},
        "defaultRun": {"latinFont": "Times New Roman", "eastAsiaFont": "宋体", "sizePt": 12},
        "styles": {
            "Normal": {
                "run": {"eastAsiaFont": "宋体", "sizePt": 12},
                "paragraph": {"lineSpacing": 1.5, "keepNext": False},
            },
            "Heading1": {
                "run": {"eastAsiaFont": "黑体", "sizePt": 16, "bold": True},
                "paragraph": {"keepNext": True},
            },
        },
        "document": {
            "type": "thesis",
            "sections": [
                {"id": "abstract", "required": True, "languages": ["zh", "en"]},
                {"id": "toc", "required": True},
                {"id": "chapter", "required": True, "repeatable": True, "min": 1},
                {"id": "references", "required": True},
            ],
        },
    }
    model = build_model(docx, profile)
    types = {item["type"] for item in model["blocks"]}
    assert "abstract_zh" in types
    assert "abstract_en" in types
    assert "chapter" in types
    assert "references" in types
    assert "toc_entry" in types
    assert model["summary"]["chapterCount"] >= 1

    diff = diff_document(docx, profile)
    failed_names = {item["name"] for item in diff["failed"]}
    assert any(name.startswith("defaultRun.eastAsiaFont") for name in failed_names)
    assert any("Heading1" in name and "eastAsiaFont" in name for name in failed_names)
    plan = build_repair_plan(diff)
    assert plan["summary"]["failed"] >= 1
    assert any(step["action"] == "set_default_run" for step in plan["repairPlan"])

    visual = visual_qa(docx, profile, None)
    issue_types = {item["type"] for item in visual["issues"]}
    assert "image_overflow" in issue_types
    assert "orphan_heading_risk" in issue_types or "orphan_heading" in issue_types
    overflow = next(item for item in visual["issues"] if item["type"] == "image_overflow")
    assert overflow.get("paragraphIndex") is not None
    plan = build_repair_plan(diff, visual)
    assert any(step.get("visual") for step in plan["repairPlan"])
    assert any(step["action"] == "scale_drawing" for step in plan["repairPlan"])

    repaired = tmp_path / "repaired.docx"
    result = apply_visual_repairs(docx, repaired, visual["issues"], force=True)
    assert result["appliedCount"] >= 1
    after = visual_qa(repaired, profile, None)
    after_types = {item["type"] for item in after["issues"]}
    assert "image_overflow" not in after_types


def test_pagination_estimate() -> None:
    model = {
        "audit": {
            "sections": [
                {
                    "page": {
                        "widthIn": 8.27,
                        "heightIn": 11.69,
                        "marginsIn": {"top": 1, "bottom": 1, "left": 1.25, "right": 1},
                    }
                }
            ],
            "defaultRun": {"sizePt": 12},
            "styles": {"Normal": {"paragraph": {"lineSpacing": 1.5}}},
        },
        "blocks": [
            {"index": 0, "section": 0, "type": "body", "textLength": 1400, "sizePt": 12, "lineSpacing": 1.5, "preview": "x" * 20},
            {"index": 1, "section": 0, "type": "chapter", "textLength": 8, "sizePt": 16, "preview": "第 3 章 系统设计", "keepNext": False},
            {"index": 2, "section": 0, "type": "body", "textLength": 80, "sizePt": 12, "lineSpacing": 1.5, "preview": "本章主要讨论"},
            {"index": 3, "section": 0, "type": "figure", "drawingHeightIn": 9.4, "drawingWidthIn": 7.2, "preview": ""},
            {"index": 4, "section": 0, "type": "figure_caption", "textLength": 12, "preview": "图 3-12 系统架构图"},
        ],
    }
    estimate = paginate(model)
    issues = detect_pagination_issues(estimate)
    types = {item["type"] for item in issues}
    assert estimate["pageCount"] >= 2
    assert "orphan_heading" in types or "heading_body_split" in types
    assert "figure_caption_split" in types


def main() -> None:
    test_compile_requirements()
    test_pagination_estimate()
    tmp = Path(__file__).with_name("_tmp_v2_fixture")
    tmp.mkdir(exist_ok=True)
    try:
        test_model_diff_visual(tmp)
    finally:
        for child in tmp.glob("*"):
            if child.is_file():
                child.unlink()
        tmp.rmdir()
    print("v2 pipeline: compile/model/diff/plan/visual/repair checks passed")


if __name__ == "__main__":
    main()
