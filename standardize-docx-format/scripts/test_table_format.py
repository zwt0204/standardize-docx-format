"""Regression tests for selected, content-safe three-line table formatting."""

import json
import zipfile
from xml.dom import minidom

import pytest

from apply_profile import apply_profile, apply_with_optional_plan
from compile_requirements import compile_text
from diff_profile import diff_document
from plan_apply import filter_profile_for_steps, select_steps
from profile_schema import validate_profile
from repair_plan import build_repair_plan
from table_format import body_tables, child, children
from test_v2_pipeline import write_docx
from validate_docx import validate_docx


def rewrite(path, transform):
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = minidom.parseString(parts["word/document.xml"])
    transform(document)
    parts["word/document.xml"] = document.toxml(encoding="utf-8")
    with zipfile.ZipFile(path, "w") as archive:
        for name, value in parts.items():
            archive.writestr(name, value)


def make_source(tmp_path, *, merged=False):
    path = tmp_path / "source.docx"
    write_docx(path)
    def add(document):
        body = document.getElementsByTagName("w:body")[0]
        for number in range(2):
            table = document.createElement("w:tbl")
            for row in range(3):
                tr = document.createElement("w:tr")
                for column in range(2):
                    xml = f'<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tcPr><w:tcBorders><w:left w:val="double"/></w:tcBorders>{"<w:vMerge/>" if merged and row == 1 else ""}</w:tcPr><w:p><w:r><w:t>表{number}行{row}列{column}</w:t></w:r></w:p></w:tc>'
                    tr.appendChild(document.importNode(minidom.parseString(xml).documentElement, True))
                table.appendChild(tr)
            # A nested table must retain its existing formatting.
            if number == 0:
                nested = table.cloneNode(True)
                cell = children(children(table, "w:tr")[2], "w:tc")[0]
                cell.insertBefore(nested, child(cell, "w:p"))
            body.insertBefore(table, child(body, "w:sectPr"))
    rewrite(path, add)
    return path


def read_document(path):
    with zipfile.ZipFile(path) as archive:
        return minidom.parseString(archive.read("word/document.xml"))


def profile():
    return {"tableRules": [{"tableIndexes": [0], "preset": "three-line", "outerSizePt": 1.5, "headerSizePt": 0.75}]}


def test_table_pipeline_and_selection(tmp_path):
    source = make_source(tmp_path)
    config = profile()
    profile_path = tmp_path / "profile.json"
    profile_path.write_text(json.dumps(config), encoding="utf-8")
    before = read_document(source)
    untouched = body_tables(before)[1].toxml()
    nested = body_tables(before)[0].getElementsByTagName("w:tbl")[0].toxml()
    differences = diff_document(source, config)
    assert not differences["ok"]
    plan = build_repair_plan(differences)
    steps = select_steps(plan, only_auto=True)
    assert steps[0]["action"] == "set_table"
    assert filter_profile_for_steps(config, []) == {}
    assert "tableRules" not in filter_profile_for_steps(config, [{"action": "set_page", "source": "page.size"}])
    plan_path = tmp_path / "plan.json"
    plan_path.write_text(json.dumps(plan), encoding="utf-8")
    output = tmp_path / "output.docx"
    apply_with_optional_plan(source, output, profile_path, False, False, plan_path=plan_path, only_auto=True, confirmed=True)
    after = read_document(output)
    tables = body_tables(after)
    assert tables[1].toxml() == untouched
    assert tables[0].getElementsByTagName("w:tbl")[0].toxml() == nested
    assert [n.firstChild.data for n in before.getElementsByTagName("w:t")] == [n.firstChild.data for n in after.getElementsByTagName("w:t")]
    rows = children(tables[0], "w:tr")
    top_cell = children(rows[0], "w:tc")[0]
    borders = child(child(top_cell, "w:tcPr"), "w:tcBorders")
    expected_order = ["w:top", "w:left", "w:start", "w:bottom", "w:right", "w:end", "w:insideH", "w:insideV"]
    assert [node.nodeName for node in children(borders)] == expected_order + ["w:tl2br", "w:tr2bl"]
    assert [node.nodeName for node in children(child(child(tables[0], "w:tblPr"), "w:tblBorders"))] == expected_order
    assert child(borders, "w:top").getAttribute("w:sz") == "12"
    assert child(borders, "w:bottom").getAttribute("w:sz") == "6"
    assert child(borders, "w:left").getAttribute("w:val") == "nil"
    assert diff_document(output, config)["ok"]
    assert validate_docx(output, profile_path)["ok"]
    def damage(document):
        cell = children(children(body_tables(document)[0], "w:tr")[1], "w:tc")[0]
        child(child(child(cell, "w:tcPr"), "w:tcBorders"), "w:right").setAttribute("w:val", "single")
    rewrite(output, damage)
    assert not diff_document(output, config)["ok"]
    assert not validate_docx(output, profile_path)["ok"]


def test_unsafe_merges_and_missing_targets_fail_without_output(tmp_path):
    source = make_source(tmp_path, merged=True)
    output = tmp_path / "bad.docx"
    with pytest.raises(ValueError, match="vertical merge"):
        apply_profile(source, output, None, False, False, profile_data=profile())
    assert not output.exists()
    config = profile()
    config["tableRules"][0]["tableIndexes"] = [99]
    with pytest.raises(ValueError, match="does not exist"):
        apply_profile(source, output, None, False, False, profile_data=config)
    assert not output.exists()


def test_multiple_header_rows_and_schema(tmp_path):
    source = make_source(tmp_path)
    config = profile()
    config["tableRules"][0].update(headerRows=2, repeatHeader=False)
    output = tmp_path / "output.docx"
    apply_profile(source, output, None, False, False, profile_data=config)
    assert diff_document(output, config)["ok"]
    for key, value in [("headerRows", 0), ("outerSizePt", 0.01), ("outerSizePt", 1.3), ("color", "bad"), ("tableIndexes", [0, 0]), ("preset", "unknown")]:
        bad = profile()
        bad["tableRules"][0][key] = value
        assert validate_profile(bad), (key, value)


def test_compiler_does_not_guess_tables_or_footer_fields():
    result = compile_text("表格采用三线表，上下线1.5磅，表头线0.75磅。")
    assert any("三线表" in value for value in result["unresolved"])
    assert any(rule["target"] == "tableRules" for rule in result["rules"])
    result = compile_text("页脚使用宋体五号。")
    assert "PAGE" not in result["profile"]["validation"].get("requiredFields", {})
    assert not result["profile"].get("headersFooters")
    result = compile_text("页脚不显示页码；目录使用自动域。")
    assert not result["profile"].get("headersFooters")
    assert result["profile"]["validation"]["requiredFields"]["TOC"] == 1
    result = compile_text("页脚留空。")
    assert result["profile"]["headersFooters"][0]["action"] == "clear"
