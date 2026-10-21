"""Explicit table layout and cell formatting without collateral content edits."""

import zipfile

import pytest

from apply_profile import apply_profile
from profile_schema import validate_profile
from table_format import body_tables, child, children, table_checks
from test_table_format import make_source, read_document, rewrite, profile


def configured():
    config = profile()
    config["tableRules"][0].update(
        paragraph={"alignment": "center", "firstLineIn": 0, "spaceAfterPt": 0, "lineSpacing": 1},
        run={"eastAsiaFont": "宋体", "latinFont": "Times New Roman", "sizePt": 10.5, "bold": False},
        allowRowSplit=False, widthIn=3, columnWidthsIn=[1, 2], alignment="center")
    return config


def check(path, config):
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    return table_checks(read_document(path), config["tableRules"], parts)[0]


def test_selected_layout_and_style_override(tmp_path):
    source = make_source(tmp_path)
    # Normal supplies exactly the properties that cell overrides must cancel.
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    from xml.dom import minidom
    styles = minidom.parseString(parts["word/styles.xml"])
    normal = next(node for node in styles.getElementsByTagName("w:style") if node.getAttribute("w:styleId") == "Normal")
    for node in list(normal.childNodes):
        if node.nodeName in {"w:pPr", "w:rPr"}:
            normal.removeChild(node)
    override = minidom.parseString('<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:ind w:firstLineChars="200"/><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>')
    for node in list(override.documentElement.childNodes):
        normal.appendChild(styles.importNode(node, True))
    parts["word/styles.xml"] = styles.toxml(encoding="utf-8")
    with zipfile.ZipFile(source, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    before = read_document(source)
    nested_before = body_tables(before)[0].getElementsByTagName("w:tbl")[0].toxml()
    other_before = body_tables(before)[1].toxml()
    config = configured()
    assert not validate_profile(config)
    output = tmp_path / "formatted.docx"
    apply_profile(source, output, None, False, False, profile_data=config)
    assert check(output, config)["ok"]
    table = body_tables(read_document(output))[0]
    assert table.getElementsByTagName("w:tbl")[0].toxml() == nested_before
    assert body_tables(read_document(output))[1].toxml() == other_before
    assert [n.getAttribute("w:w") for n in children(child(table, "w:tblGrid"), "w:gridCol")] == ["1440", "2880"]
    assert child(child(table, "w:tblPr"), "w:tblW").getAttribute("w:w") == "4320"
    assert child(child(children(table, "w:tr")[0], "w:trPr"), "w:cantSplit").getAttribute("w:val") == "true"
    # A damaged value must be detected independently of the writer.
    def damage(document):
        cell = children(children(body_tables(document)[0], "w:tr")[0], "w:tc")[0]
        paragraph = child(cell, "w:p")
        child(child(paragraph, "w:pPr"), "w:ind").setAttribute("w:firstLineChars", "200")
        run = child(paragraph, "w:r")
        child(child(run, "w:rPr"), "w:sz").setAttribute("w:val", "28")
    rewrite(output, damage)
    failures = check(output, config)
    assert not failures["ok"]
    assert any("firstLineIn" in issue for issue in failures["actual"])
    assert any("sizePt" in issue for issue in failures["actual"])


@pytest.mark.parametrize("property_name", ["width", "grid", "split", "alignment", "inherited_font"])
def test_layout_damage_is_detected(tmp_path, property_name):
    source = make_source(tmp_path)
    config = configured()
    output = tmp_path / "formatted.docx"
    apply_profile(source, output, None, False, False, profile_data=config)
    def damage(document):
        table = body_tables(document)[0]
        if property_name == "width":
            child(child(table, "w:tblPr"), "w:tblW").setAttribute("w:w", "999")
        elif property_name == "grid":
            children(child(table, "w:tblGrid"), "w:gridCol")[0].setAttribute("w:w", "999")
        elif property_name == "split":
            child(child(children(table, "w:tr")[0], "w:trPr"), "w:cantSplit").setAttribute("w:val", "false")
        elif property_name == "alignment":
            child(child(table, "w:tblPr"), "w:jc").setAttribute("w:val", "right")
        else:
            paragraph = child(children(children(table, "w:tr")[0], "w:tc")[0], "w:p")
            run = child(paragraph, "w:r")
            fonts = child(child(run, "w:rPr"), "w:rFonts")
            fonts.removeAttribute("w:eastAsia")
    rewrite(output, damage)
    assert not check(output, config)["ok"]


def test_grid_span_and_rejection_without_output(tmp_path):
    source = make_source(tmp_path)
    def merge(document):
        table = body_tables(document)[0]
        row = children(table, "w:tr")[0]
        cells = children(row, "w:tc")
        node = document.createElement("w:gridSpan")
        node.setAttribute("w:val", "2")
        child(cells[0], "w:tcPr").appendChild(node)
        row.removeChild(cells[1])
    rewrite(source, merge)
    config = configured()
    output = tmp_path / "span.docx"
    apply_profile(source, output, None, False, False, profile_data=config)
    cell = children(children(body_tables(read_document(output))[0], "w:tr")[0], "w:tc")[0]
    assert child(child(cell, "w:tcPr"), "w:tcW").getAttribute("w:w") == "4320"
    assert check(output, config)["ok"]
    config["tableRules"][0].update(widthIn=4, columnWidthsIn=[1, 1, 2])
    rejected = tmp_path / "rejected.docx"
    with pytest.raises(ValueError, match="does not match"):
        apply_profile(source, rejected, None, False, False, profile_data=config)
    assert not rejected.exists()


@pytest.mark.parametrize("changes", [
    {"widthIn": 5}, {"columnWidthsIn": [0, 3]}, {"paragraph": {"alignment": "distribute"}},
    {"paragraph": {"firstLineIn": 0, "hangingIn": 1}}, {"run": {"sizePt": 10.3}},
    {"run": {"color": "bad"}}, {"run": {"updateTheme": True}},
    {"paragraph": {"unknown": 0}}, {"paragraph": {"lineSpacing": 0}},
])
def test_explicit_layout_schema_rejects_unsafe_or_unsupported(changes):
    config = configured()
    config["tableRules"][0].update(changes)
    assert validate_profile(config)
