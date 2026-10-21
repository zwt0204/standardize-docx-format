#!/usr/bin/env python3
"""Behavioral regressions for measurable header/footer compliance."""
from __future__ import annotations

import tempfile
import unittest
import zipfile
from xml.dom import minidom
from pathlib import Path

from apply_profile import apply_profile
from audit_docx import audit, parse_xml, marginal_effective_props, marginal_style_context
from apply_profile import apply_paragraph_format
from diff_profile import diff_headers
from test_v2_pipeline import write_docx
from validate_docx import validate_profile_expectations


class MarginalFormatTests(unittest.TestCase):
    def rewrite(self, transform):
        with zipfile.ZipFile(self.output) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        transform(parts)
        with zipfile.ZipFile(self.output, "w") as archive:
            for name, content in parts.items():
                archive.writestr(name, content)

    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.source = Path(self.directory.name) / "source.docx"
        self.output = Path(self.directory.name) / "output.docx"
        write_docx(self.source)

    def apply(self, rule: dict) -> dict:
        apply_profile(self.source, self.output, None, False,
                      profile_data={"headersFooters": [rule]}, skip_schema=True)
        return audit(self.output)

    def assertCompliance(self, report: dict, rule: dict, ok: bool) -> None:
        findings: list = []
        diff_headers(findings, report, {"headersFooters": [rule]})
        checks: list = []
        validate_profile_expectations(checks, report, {"headersFooters": [rule]})
        self.assertEqual([item["ok"] for item in findings], [ok], findings)
        self.assertEqual([item["ok"] for item in checks], [ok], checks)

    def test_text_format_border_and_nested_paragraph(self) -> None:
        rule = {"section": 0, "kind": "header", "text": "论文标题",
                "paragraph": {"alignment": "right", "spaceAfterPt": 3, "firstLineChars": 1},
                "run": {"eastAsiaFont": "宋体", "sizePt": 10, "color": "#123456"},
                "bottomBorder": {"size": 8}}
        report = self.apply(rule)
        self.assertCompliance(report, rule, True)
        for bad in ({**rule, "text": "错误标题"},
                    {**rule, "run": {"sizePt": 12}},
                    {**rule, "bottomBorder": {"size": 10}},
                    {**rule, "paragraph": {"alignment": "left"}}):
            self.assertCompliance(report, bad, False)

    def test_field_content_not_cached_result(self) -> None:
        rule = {"section": 0, "kind": "footer", "segments": [
            {"text": "第 "}, {"field": "PAGE", "result": "1", "run": {"sizePt": 9}}, {"text": " 页"}]}
        report = self.apply(rule)
        changed_cache = {**rule, "segments": [
            {"text": "第 "}, {"field": "PAGE", "result": "99", "run": {"sizePt": 9}}, {"text": " 页"}]}
        self.assertCompliance(report, changed_cache, True)
        self.assertCompliance(report, {**rule, "segments": [{"field": "NUMPAGES"}]}, False)
        self.assertCompliance(report, {**rule, "action": "clear"}, False)

    def test_clear_even_and_first_flags(self) -> None:
        for reference_type in ("even", "first"):
            rule = {"section": 0, "kind": "header", "type": reference_type, "action": "clear"}
            report = self.apply(rule)
            self.assertCompliance(report, rule, True)
            if reference_type == "first":
                report["sections"][0]["differentFirstPage"] = False
            else:
                report["marginals"]["evenAndOddHeaders"] = False
            self.assertCompliance(report, rule, False)
            self.output.unlink()

    def test_missing_part_and_simple_field(self) -> None:
        rule = {"section": 0, "kind": "footer", "segments": [{"field": "PAGE", "run": {"sizePt": 9}}]}
        report = self.apply(rule)
        part = report["marginals"]["parts"][0]
        part_name = part["part"]
        with zipfile.ZipFile(self.output) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        parts[part_name] = b'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>99</w:t></w:r></w:fldSimple></w:p></w:ftr>'
        custom_name = "word/custom-footer.xml"
        parts[custom_name] = parts.pop(part_name)
        parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(
            part_name.removeprefix("word/").encode(), b"custom-footer.xml")
        with zipfile.ZipFile(self.output, "w") as archive:
            for name, content in parts.items():
                archive.writestr(name, content)
        custom_report = audit(self.output)
        self.assertCompliance(custom_report, rule, True)
        self.assertEqual(custom_report["fields"]["kinds"]["PAGE"], 1)
        report["marginals"]["parts"] = []
        self.assertCompliance(report, rule, False)

    def test_zero_indent_overrides_inherited_character_indents(self) -> None:
        rule = {"section": 0, "kind": "footer", "text": "page", "paragraph": {"firstLineIn": 0}}
        self.apply(rule)
        def change(parts):
            parts["word/styles.xml"] = parts["word/styles.xml"].replace(
                b'<w:pPr>', b'<w:pPr><w:ind w:firstLineChars="200" w:hangingChars="100"/>', 1)
        self.rewrite(change)
        report = audit(self.output)
        self.assertCompliance(report, rule, True)
        props = report["marginals"]["parts"][0]["paragraphDetails"][0]["paragraph"]
        self.assertEqual(props["firstLineIn"], 0)
        self.assertEqual(props["hangingIn"], 0)

    def test_character_precedence_is_evaluated_after_inheritance(self) -> None:
        ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        styles = ('<w:styles xmlns:w="' + ns + '"><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                  '<w:pPr><w:ind w:firstLineChars="200"/></w:pPr></w:style></w:styles>')
        context = marginal_style_context({"word/styles.xml": styles.encode()})
        paragraph = parse_xml(('<w:p xmlns:w="' + ns + '"><w:pPr><w:ind w:firstLine="0"/></w:pPr></w:p>').encode()).documentElement
        props = marginal_effective_props(paragraph, None, context)[0]
        self.assertIsNone(props["firstLineIn"])
        self.assertEqual(props["firstLineChars"], 2)
        p_pr = paragraph.getElementsByTagName('w:pPr')[0]
        apply_paragraph_format(p_pr, {"firstLineIn": 0}, 12)
        self.assertEqual(marginal_effective_props(paragraph, None, context)[0]["firstLineIn"], 0)
        paragraph.ownerDocument._format_styles_document = context[0]
        with self.assertRaisesRegex(ValueError, "Cannot safely set nonzero"):
            apply_paragraph_format(p_pr, {"hangingIn": 0.25}, 12)
        clean = parse_xml(('<w:p xmlns:w="' + ns + '"><w:pPr/></w:p>').encode()).documentElement
        p_pr = clean.getElementsByTagName('w:pPr')[0]
        clean_context = (None, {}, "Normal")
        apply_paragraph_format(p_pr, {"hangingIn": 0.25}, 12)
        props = marginal_effective_props(clean, None, clean_context)[0]
        self.assertEqual(props["firstLineIn"], 0)
        self.assertEqual(props["hangingIn"], 0.25)
        apply_paragraph_format(p_pr, {"firstLineChars": 2}, 12)
        props = marginal_effective_props(clean, None, clean_context)[0]
        self.assertEqual(props["hangingIn"], 0)
        self.assertAlmostEqual(props["firstLineIn"], 1 / 3, places=4)

    def test_nonzero_physical_indent_is_not_approximated_as_characters(self) -> None:
        rule = {"section": 0, "kind": "header", "text": "header", "paragraph": {"firstLineIn": 0.25}}
        report = self.apply(rule)
        self.assertCompliance(report, rule, True)
        with zipfile.ZipFile(self.output) as archive:
            xml = archive.read("word/header1.xml")
        self.assertIn(b'w:firstLine="360"', xml)
        self.assertNotIn(b'firstLineChars', xml)
        self.assertNotIn(b'hanging', xml)

    def test_nonzero_indent_rejects_inherited_character_or_opposite_indent(self) -> None:
        for inherited in ('w:firstLineChars="200"', 'w:hanging="120"', 'w:firstLineChars="0"'):
            write_docx(self.source)
            with zipfile.ZipFile(self.source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            parts["word/styles.xml"] = parts["word/styles.xml"].replace(
                b'<w:pPr>', ('<w:pPr><w:ind ' + inherited + '/>').encode(), 1)
            with zipfile.ZipFile(self.source, "w") as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
            rule = {"section": 0, "kind": "footer", "text": "page", "paragraph": {"firstLineIn": 0.25}}
            with self.assertRaisesRegex(ValueError, "Cannot safely set nonzero physical indentation"):
                self.apply(rule)
            self.assertFalse(self.output.exists())

    def test_zero_indent_does_not_hide_competing_character_attributes(self) -> None:
        rule = {"section": 0, "kind": "footer", "text": "page", "paragraph": {"firstLineIn": 0}}
        self.apply(rule)
        def change(parts):
            parts["word/footer1.xml"] = parts["word/footer1.xml"].replace(
                b'w:firstLineChars="0"', b'w:firstLineChars="200"')
        self.rewrite(change)
        self.assertCompliance(audit(self.output), rule, False)

    def test_later_parent_style_is_checked_for_conflicts(self) -> None:
        profile = {"styles": {
            "Child": {"basedOn": "Parent", "paragraph": {"firstLineIn": 0.25}},
            "Parent": {"paragraph": {"hangingIn": 0.125}}}}
        with self.assertRaisesRegex(ValueError, "Cannot safely set nonzero physical indentation"):
            apply_profile(self.source, self.output, None, False, profile_data=profile, skip_schema=True)
        self.assertFalse(self.output.exists())

    def test_markup_compatibility_reads_only_selected_ordinary_branch(self) -> None:
        rule = {"section": 0, "kind": "footer", "segments": [{"field": "PAGE"}]}
        self.apply(rule)
        def change(parts):
            parts["word/footer1.xml"] = b'''<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:known="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:unknown="urn:unknown"><mc:AlternateContent><mc:Choice Requires="unknown"><w:p><w:r><w:t>wrong</w:t></w:r></w:p></mc:Choice><mc:Choice Requires="known"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:ftr>'''
        self.rewrite(change)
        before = self.output.read_bytes()
        report = audit(self.output)
        self.assertCompliance(report, rule, True)
        self.assertEqual(report["marginals"]["parts"][0]["paragraphs"], 1)
        self.assertEqual(report["fields"]["kinds"]["PAGE"], 1)
        self.assertEqual(self.output.read_bytes(), before)

    def test_markup_compatibility_floating_page_number_is_unsupported(self) -> None:
        rule = {"section": 0, "kind": "footer", "segments": [{"field": "PAGE"}]}
        self.apply(rule)
        def change(parts):
            parts["word/footer1.xml"] = b'''<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:wsp><w:txbxContent><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:txbxContent></wps:wsp></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><w:txbxContent><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:txbxContent></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p></w:ftr>'''
        self.rewrite(change)
        report = audit(self.output)
        part = report["marginals"]["parts"][0]
        self.assertEqual(report["fields"]["kinds"]["PAGE"], 1)
        self.assertEqual(part["paragraphs"], 2)
        self.assertFalse(part["hasObjects"])
        self.assertTrue(part["unsupportedLayouts"])
        self.assertCompliance(report, rule, False)

    def test_inherit_preserves_official_template(self) -> None:
        rule = {"section": 0, "kind": "header", "text": "官方模板"}
        report = self.apply(rule)
        inherited = Path(self.directory.name) / "inherit.docx"
        apply_profile(self.output, inherited, None, False, profile_data={"headersFooters": [
            {"section": 0, "kind": "header", "action": "inherit"}]}, skip_schema=True)
        self.assertEqual(audit(inherited)["sections"][0]["headerReferences"], report["sections"][0]["headerReferences"])

    def test_additional_visible_content_is_not_compliant(self) -> None:
        rule = {"section": 0, "kind": "header", "text": "论文标题"}
        self.apply(rule)
        original = self.output.read_bytes()
        for content in ('<w:drawing/>', '<w:sym w:font="Wingdings" w:char="F0FC"/>', '<w:tab/>'):
            self.output.write_bytes(original)
            def change(parts):
                parts["word/header1.xml"] = parts["word/header1.xml"].replace(b'</w:p>', ('<w:r>' + content + '</w:r></w:p>').encode())
            self.rewrite(change)
            self.assertCompliance(audit(self.output), rule, False)

    def test_inherited_run_and_paragraph_properties(self) -> None:
        rule = {"section": 0, "kind": "header", "text": "论文标题", "run": {"sizePt": 10, "eastAsiaFont": "宋体", "bold": True}, "paragraph": {"alignment": "right"}}
        self.apply(rule)
        def change(parts):
            document = minidom.parseString(parts["word/header1.xml"])
            run = document.getElementsByTagName("w:r")[0]
            run.removeChild(run.getElementsByTagName("w:rPr")[0])
            run.insertBefore(document.importNode(minidom.parseString('<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rStyle w:val="HeaderText"/></w:rPr>').documentElement, True), run.firstChild)
            paragraph = document.getElementsByTagName("w:p")[0]
            paragraph.removeChild(paragraph.getElementsByTagName("w:pPr")[0])
            paragraph.insertBefore(document.importNode(minidom.parseString('<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pStyle w:val="HeaderChild"/></w:pPr>').documentElement, True), paragraph.firstChild)
            parts["word/header1.xml"] = document.toxml(encoding="utf-8")
            styles = minidom.parseString(parts["word/styles.xml"])
            styles.getElementsByTagName("w:sz")[0].setAttribute("w:val", "20")
            for xml in (
                '<w:style w:type="paragraph" w:styleId="HeaderBase"><w:pPr><w:jc w:val="right"/></w:pPr><w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr></w:style>',
                '<w:style w:type="paragraph" w:styleId="HeaderChild"><w:basedOn w:val="HeaderBase"/></w:style>',
                '<w:style w:type="character" w:styleId="HeaderText"><w:rPr><w:b/></w:rPr></w:style>'):
                element = minidom.parseString(xml.replace('<w:style ', '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ', 1)).documentElement
                styles.documentElement.appendChild(styles.importNode(element, True))
            parts["word/styles.xml"] = styles.toxml(encoding="utf-8")
        self.rewrite(change)
        self.assertCompliance(audit(self.output), rule, True)
        def override(parts):
            parts["word/header1.xml"] = parts["word/header1.xml"].replace(b'<w:rStyle w:val="HeaderText"/>', b'<w:rStyle w:val="HeaderText"/><w:sz w:val="24"/>')
        self.rewrite(override)
        self.assertCompliance(audit(self.output), rule, False)


if __name__ == "__main__":
    unittest.main()
