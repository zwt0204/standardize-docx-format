"""Conservative table rules from explicit specifications or template borders."""
import re

from table_format import body_tables, child, children, table_checks, has_compatibility_tables


def compile_table_clause(line):
    """Only compile explicit zero-based body indexes and measured borders."""
    locator = re.search(r'(?:正文表格索引|body table indexes?)\s*[:：=]?\s*(\d+(?:\s*[,，、]\s*\d+)*)', line, re.I)
    header = re.search(r'(?:表头行数|header rows?)\s*[:：=]?\s*(\d+)', line, re.I)
    outer = re.search(r'(?:上下线|outer)\s*[:：=]?\s*([\d.]+)\s*(?:磅|pt)', line, re.I)
    middle = re.search(r'(?:表头线|header border)\s*[:：=]?\s*([\d.]+)\s*(?:磅|pt)', line, re.I)
    if not all((locator, header, outer, middle)):
        return None
    return {'tableIndexes': [int(n) for n in re.split(r'\s*[,，、]\s*', locator[1])],
            'preset': 'three-line', 'headerRows': int(header[1]),
            'outerSizePt': float(outer[1]), 'headerSizePt': float(middle[1])}


def infer_template_tables(document):
    """Recognize fully explicit borders; never infer through table styles."""
    if has_compatibility_tables(document):
        return [], ["模板兼容分支包含表格，正文表格索引不明确；须确认保留模板的定位方式，不自动提取 tableRules。"]
    rules, unresolved = [], []
    for index, table in enumerate(body_tables(document)):
        rows = children(table, 'w:tr')
        props = child(table, 'w:tblPr')
        borders = child(props, 'w:tblBorders')
        top, bottom = child(borders, 'w:top'), child(borders, 'w:bottom')
        headers = 0
        for row in rows:
            marker = child(child(row, 'w:trPr'), 'w:tblHeader')
            if marker is None or marker.getAttribute('w:val') in {'0', 'false', 'off'}:
                break
            headers += 1
        candidate = None
        if top is not None and bottom is not None and 0 < headers < len(rows):
            cells = children(rows[headers - 1], 'w:tc')
            middle = child(child(child(cells[0], 'w:tcPr'), 'w:tcBorders'), 'w:bottom') if cells else None
            try:
                if middle is not None and all(n.getAttribute('w:val') == 'single' for n in (top, bottom, middle)) and top.getAttribute('w:sz') == bottom.getAttribute('w:sz') and len({n.getAttribute('w:color').upper() for n in (top, bottom, middle)}) == 1:
                    candidate = {'tableIndexes': [index], 'preset': 'three-line', 'headerRows': headers,
                                 'outerSizePt': int(top.getAttribute('w:sz')) / 8,
                                 'headerSizePt': int(middle.getAttribute('w:sz')) / 8,
                                 'color': top.getAttribute('w:color')}
                    # Check only this table, avoiding accidental template-index mapping.
                    probe = dict(candidate, tableIndexes=[0])
                    holder = document.implementation.createDocument('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:body', None)
                    holder.documentElement.appendChild(holder.importNode(table, True))
                    if not table_checks(holder, [probe])[0]['ok']:
                        candidate = None
            except (ValueError, IndexError):
                candidate = None
        if candidate:
            rules.append(candidate)
            unresolved.append(f'模板三线表索引 {index} 已提取；应用前确认目标文档对应的正文表格索引，不包含嵌套表。')
        else:
            unresolved.append(f'模板正文表格索引 {index} 未能保守识别为显式三线表；请确认表头、边框及目标映射。')
    return rules, unresolved
