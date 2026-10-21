"""Behavioral coverage for native flowcharts and non-destructive DOCX insertion."""

import copy
import json
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.dom import minidom

import pytest

from audit_docx import audit
from flowchart import A, WPG, WPS, add_flowchart, blank_package, grid_positions, validate_spec


def example():
    return json.loads((Path(__file__).parent.parent / "references/flowchart.example.json").read_text(encoding="utf-8"))


def read_parts(path):
    with zipfile.ZipFile(path) as archive:
        assert archive.testzip() is None
        return {name: archive.read(name) for name in archive.namelist()}


def document(path):
    return minidom.parseString(read_parts(path)["word/document.xml"])


def test_branch_return_and_native_connections(tmp_path):
    output = tmp_path / "chart.docx"
    result = add_flowchart(example(), output)
    dom = document(output)
    assert result["summary"]["nodes"] == 4
    assert result["summary"]["edges"] == 4
    assert result["visualConfirmation"] == "pending"
    assert len(dom.getElementsByTagNameNS(WPG, "wgp")) == 1
    assert len(dom.getElementsByTagNameNS(WPS, "cNvCnPr")) == 4
    assert len(dom.getElementsByTagNameNS(WPS, "txbx")) == 6
    # The flowchart has Word text and shapes; it is not a raster image.
    assert not any(name.startswith("word/media/") for name in read_parts(output))
    ids = [n.getAttribute("id") for n in dom.getElementsByTagNameNS(WPS, "cNvPr")]
    assert len(ids) == len(set(ids))
    node_ids = {n.getAttribute("name"): n.getAttribute("id")
                for n in dom.getElementsByTagNameNS(WPS, "cNvPr") if n.getAttribute("name") in {"submit", "review", "archive", "revise"}}
    connectors = dom.getElementsByTagNameNS(WPS, "cNvCnPr")
    for edge, connector in zip(example()["edges"], connectors):
        start = connector.getElementsByTagNameNS(A, "stCxn")[0]
        end = connector.getElementsByTagNameNS(A, "endCxn")[0]
        assert start.getAttribute("id") == node_ids[edge["from"]]
        assert end.getAttribute("id") == node_ids[edge["to"]]
        assert int(start.getAttribute("idx")) in range(4)
    assert all(s.lastChild.nodeName == "wps:bodyPr" for s in dom.getElementsByTagNameNS(WPS, "wsp"))
    assert audit(output)["content"]["drawings"] == 1


def make_source(path):
    parts = blank_package()
    xml = parts["word/document.xml"].decode().replace("<w:sectPr>", '<w:p><w:r><w:t>插入位置</w:t></w:r></w:p><w:p><w:r><w:t>原文保留</w:t></w:r></w:p><w:sectPr>')
    # Compatibility branches and unknown parts must survive the insertion.
    xml = xml.replace("<w:body>", '<w:body xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown"><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>原始分支</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>兼容分支</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>')
    parts["word/document.xml"] = xml.encode()
    parts["customXml/item1.xml"] = b'<preserve untouched="yes"/>'
    parts["word/media/user.bin"] = b"user-data"
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def test_insertion_preserves_package_and_content_and_allocates_ids(tmp_path):
    source, output = tmp_path / "source.docx", tmp_path / "out.docx"
    make_source(source)
    original = source.read_bytes()
    add_flowchart(example(), output, source=source, after_text="插入位置")
    before, after = read_parts(source), read_parts(output)
    assert source.read_bytes() == original
    assert {k: v for k, v in before.items() if k != "word/document.xml"} == {k: v for k, v in after.items() if k != "word/document.xml"}
    dom = document(output)
    assert len(dom.getElementsByTagName("mc:Choice")) == 1
    assert len(dom.getElementsByTagName("mc:Fallback")) == 1
    body = dom.getElementsByTagName("w:body")[0]
    children = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE]
    assert children[-1].nodeName == "w:sectPr"
    assert children[2].getElementsByTagName("w:drawing")
    second = tmp_path / "second.docx"
    add_flowchart(example(), second, source=output)
    ids = [n.getAttribute("id") for n in document(second).getElementsByTagName("*") if n.localName in {"docPr", "cNvPr"}]
    assert len(ids) == len(set(ids))


@pytest.mark.parametrize("mutate", [
    lambda s: s["nodes"].append(copy.deepcopy(s["nodes"][0])),
    lambda s: s["edges"][0].update(to="missing"),
    lambda s: s["nodes"][0].update(text="bad\x01text"),
    lambda s: s["nodes"][0].update(type="smartart"),
    lambda s: s["nodes"][0].update(row=0, column=0),
    lambda s: s.update(layout={"nodeWidthIn": float("nan")}),
    lambda s: s.update(layout={"fontSizePt": True}),
    lambda s: s["edges"].append(copy.deepcopy(s["edges"][0])),
    lambda s: s.update(unknown=True),
    lambda s: s["nodes"][0].update(text="非常长的文字" * 30),
])
def test_invalid_specs_do_not_replace_existing_output(tmp_path, mutate):
    spec = example()
    mutate(spec)
    output = tmp_path / "existing.docx"
    output.write_bytes(b"keep me")
    with pytest.raises(ValueError):
        add_flowchart(spec, output, force=True)
    assert output.read_bytes() == b"keep me"


def test_output_and_locator_guards(tmp_path):
    source = tmp_path / "source.docx"
    make_source(source)
    original = source.read_bytes()
    with pytest.raises(ValueError, match="in-place"):
        add_flowchart(example(), source, source=source, force=True)
    output = tmp_path / "out.docx"
    with pytest.raises(ValueError, match="matched 0"):
        add_flowchart(example(), output, source=source, after_text="不存在")
    assert not output.exists()
    add_flowchart(example(), output)
    with pytest.raises(FileExistsError):
        add_flowchart(example(), output)
    assert source.read_bytes() == original


def test_sequential_explicit_self_loop_and_size_guard(tmp_path):
    spec = {"nodes": [{"id": "start", "text": "开始", "type": "start"},
                      {"id": "end", "text": "结束", "type": "end"}],
            "edges": [{"from": "start", "to": "end"}, {"from": "end", "to": "end"}]}
    add_flowchart(spec, tmp_path / "self.docx")
    for i, node in enumerate(spec["nodes"]):
        node.update(row=i, column=0)
    validate_spec(spec)
    assert grid_positions(spec) == {"start": (0, 0), "end": (1, 0)}
    add_flowchart(spec, tmp_path / "explicit.docx")
    spec["nodes"][1]["row"] = 59
    with pytest.raises(ValueError, match="exceeds section"):
        add_flowchart(spec, tmp_path / "oversized.docx")


def test_branch_merge_and_skipped_rank_routing(tmp_path):
    spec = {"nodes": [{"id": "a", "text": "分支", "type": "decision"},
                      {"id": "b", "text": "路径一"}, {"id": "c", "text": "路径二"},
                      {"id": "d", "text": "合流", "type": "end"}],
            "edges": [{"from": "a", "to": "b", "label": "是"},
                      {"from": "a", "to": "c", "label": "否"},
                      {"from": "b", "to": "d"}, {"from": "c", "to": "d"}]}
    result = add_flowchart(spec, tmp_path / "merge.docx")
    assert result["summary"]["grid"]["b"][0] == result["summary"]["grid"]["c"][0]
    assert result["summary"]["grid"]["d"][0] > result["summary"]["grid"]["b"][0]
    # A long edge must go around the intervening node, not through it.
    spec = {"nodes": [{"id": "a", "text": "A"}, {"id": "b", "text": "B"}, {"id": "c", "text": "C"}],
            "edges": [{"from": "a", "to": "b"}, {"from": "b", "to": "c"}, {"from": "a", "to": "c"}]}
    path = tmp_path / "skip.docx"
    add_flowchart(spec, path)
    conn = document(path).getElementsByTagNameNS(WPS, "cNvCnPr")[-1]
    assert conn.getElementsByTagNameNS(A, "stCxn")[0].getAttribute("idx") == "3"


def test_escaped_text_round_trip(tmp_path):
    text = 'A&B <输入> "值"\n下一步'
    spec = {"nodes": [{"id": "a", "text": text}], "edges": [],
            "layout": {"nodeWidthIn": 3, "nodeHeightIn": 1}}
    output = tmp_path / "escaped.docx"
    add_flowchart(spec, output)
    lines = [n.firstChild.data for n in document(output).getElementsByTagName("w:t")]
    assert "\n".join(lines) == text


@pytest.mark.parametrize("change, match", [
    (lambda xml: xml.replace("<w:sectPr>", '<w:p><w:r><w:t>插入位置</w:t></w:r></w:p><w:sectPr>'), "matched 2"),
    (lambda xml: xml.replace("</w:sectPr>", '<w:cols w:num="2"/></w:sectPr>'), "Multi-column"),
])
def test_unsafe_insertion_contexts(tmp_path, change, match):
    source, output = tmp_path / "source.docx", tmp_path / "out.docx"
    make_source(source)
    parts = read_parts(source)
    parts["word/document.xml"] = change(parts["word/document.xml"].decode()).encode()
    with zipfile.ZipFile(source, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    with pytest.raises(ValueError, match=match):
        add_flowchart(example(), output, source=source, after_text="插入位置")
    assert not output.exists()


def test_cli_and_spec_protection(tmp_path):
    cli = Path(__file__).with_name("standardize.py")
    spec = tmp_path / "spec.json"
    spec.write_text(json.dumps(example()), encoding="utf-8")
    output = tmp_path / "chart.docx"
    result = subprocess.run([sys.executable, str(cli), "flowchart", "--spec", str(spec), "--output", str(output)], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["summary"]["edges"] == 4
    before = spec.read_bytes()
    result = subprocess.run([sys.executable, str(cli), "flowchart", "--spec", str(spec), "--output", str(spec), "--force"], capture_output=True, text=True)
    assert result.returncode == 1
    assert spec.read_bytes() == before
