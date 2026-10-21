"""Compiler/template routing and header linkage impact regressions."""
from compile_requirements import compile_text
from diff_profile import diff_headers
from repair_plan import build_repair_plan
from table_inference import infer_template_tables
from apply_profile import apply_profile
from test_table_format import make_source, profile, read_document


def test_explicit_and_ambiguous_table_spec():
    result = compile_text('正文宋体小四。\n正文表格索引0,2采用三线表，表头行数1，上下线1.5磅，表头线0.75磅')
    assert result['profile']['tableRules'] == [dict(tableIndexes=[0, 2], preset='three-line', headerRows=1, outerSizePt=1.5, headerSizePt=.75)]
    assert not any('三线表' in item for item in result['unresolved'])
    result = compile_text('所有表格均采用三线表，上下线1.5磅，表头线0.75磅')
    assert not result['profile'].get('tableRules')
    assert any('三线表' in item for item in result['unresolved'])


def test_template_only_infers_explicit_tables(tmp_path):
    source = make_source(tmp_path)
    rules, unresolved = infer_template_tables(read_document(source))
    assert not rules and len(unresolved) == 2
    output = tmp_path / 'formatted.docx'
    apply_profile(source, output, None, False, False, profile_data=profile())
    rules, unresolved = infer_template_tables(read_document(output))
    assert len(rules) == 1 and rules[0]['tableIndexes'] == [0]
    assert rules[0]['outerSizePt'] == 1.5 and rules[0]['headerSizePt'] == .75
    assert any('确认目标' in item for item in unresolved)


def test_plan_reports_linked_sections_without_freezing():
    sections = [{'headerReferences': []}, {'headerReferences': []},
                {'headerReferences': [{'type': 'default', 'target': 'word/header2.xml'}]}, {'headerReferences': []}]
    findings = []
    diff_headers(findings, {'sections': sections}, {'headersFooters': [{'section': 0, 'kind': 'header', 'text': 'Title'}]})
    assert findings[0]['impact']['inheritingSections'] == [1]
    plan = build_repair_plan({'failed': findings})
    assert plan['repairPlan'][0]['impact']['inheritingSections'] == [1]
    assert not any(step['action'] == 'freeze_section' for step in plan['repairPlan'])


def test_compatibility_tables_refuse_ambiguous_indexes(tmp_path):
    import pytest
    from table_format import body_tables
    from test_table_format import rewrite
    source = make_source(tmp_path)
    def wrap(document):
        table = body_tables(document)[0]
        mc = document.createElementNS('http://schemas.openxmlformats.org/markup-compatibility/2006', 'mc:AlternateContent')
        mc.setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        choice, fallback = document.createElement('mc:Choice'), document.createElement('mc:Fallback')
        choice.setAttribute('Requires', 'w')
        table.parentNode.insertBefore(mc, table)
        choice.appendChild(table)
        fallback.appendChild(table.cloneNode(True))
        mc.appendChild(choice)
        mc.appendChild(fallback)
    rewrite(source, wrap)
    rules, unresolved = infer_template_tables(read_document(source))
    assert not rules and any('兼容分支' in item for item in unresolved)
    from analyze_template import analyze_template
    analysis = analyze_template(source)
    assert not analysis['profile'].get('tableRules')
    assert any('兼容分支' in item for item in analysis['unresolved'])
    output = tmp_path / 'unsafe.docx'
    with pytest.raises(ValueError, match='markup compatibility'):
        apply_profile(source, output, None, False, False, profile_data=profile())
    assert not output.exists()
    # Unrelated style-only work must remain usable and preserve both branches.
    apply_profile(source, output, None, False, False, profile_data={'defaultRun': {'sizePt': 12}})
    assert len(read_document(output).getElementsByTagNameNS('http://schemas.openxmlformats.org/markup-compatibility/2006', 'AlternateContent')) == 1


def test_table_border_points_do_not_become_body_font_size():
    result = compile_text('正文表格索引0采用三线表，表头行数1，上下线1.5磅，表头线0.75磅')
    assert result['profile']['defaultRun']['sizePt'] == 12
    assert result['profile']['styles']['Normal']['run']['sizePt'] == 12
    assert any('正文字号未明确' in item for item in result['unresolved'])
    result = compile_text('正文宋体小四。\n正文表格索引0采用三线表，表头行数1，上下线1.5磅，表头线0.75磅')
    assert result['profile']['defaultRun']['sizePt'] == 12


def test_unsupported_marginal_is_inspect_not_auto_rebuild():
    plan = build_repair_plan({'failed': [{'category': 'headersFooters', 'name': 'headersFooters[0]',
        'actual': {'errors': ['unsupported: floating text box']},
        'expected': {'section': 0, 'kind': 'header', 'text': 'Title'}}]})
    assert plan['repairPlan'][0]['action'] == 'inspect'
    assert not plan['repairPlan'][0]['autoApplyable']
