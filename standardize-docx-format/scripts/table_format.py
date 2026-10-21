"""Explicit, content-preserving three-line borders for selected body tables."""

from __future__ import annotations

from xml.dom import Node


ORDERS = {
    "w:tblPr": "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange",
    "w:trPr": "cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden ins del trPrChange",
    "w:tcPr": "cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge tcPrChange",
}
EDGES = "top left start bottom right end insideH insideV".split()


def children(node, tag=None):
    return [c for c in node.childNodes if c.nodeType == Node.ELEMENT_NODE and (tag is None or c.nodeName == tag)]


def child(node, tag):
    return next(iter(children(node, tag)), None) if node is not None else None


def ensure(node, tag):
    found = child(node, tag)
    if found is not None:
        return found
    found = node.ownerDocument.createElement(tag)
    order = ["w:" + value for value in ORDERS.get(node.nodeName, "").split()]
    if tag in {"w:tblPr", "w:trPr", "w:tcPr"}:
        node.insertBefore(found, node.firstChild)
    elif tag in order:
        before = next((c for c in children(node) if c.nodeName in order and order.index(c.nodeName) > order.index(tag)), None)
        node.insertBefore(found, before)
    else:
        node.appendChild(found)
    return found


def has_compatibility_tables(document):
    return any(node.getElementsByTagName("w:tbl") for node in document.getElementsByTagNameNS(
        "http://schemas.openxmlformats.org/markup-compatibility/2006", "AlternateContent"))


def body_tables(document):
    """Index outer body tables only; never touch nested or text-box tables."""
    if has_compatibility_tables(document):
        raise ValueError("Table indexes are ambiguous inside markup compatibility branches; use a template-preserving locator edit")
    result = []
    for table in document.getElementsByTagName("w:tbl"):
        parent = table.parentNode
        while parent is not None and parent.nodeName not in {"w:tbl", "w:txbxContent", "w:body"}:
            parent = parent.parentNode
        if parent is not None and parent.nodeName == "w:body":
            result.append(table)
    return result


def border_values(size, color):
    return {"w:val": "single", "w:sz": str(round(size * 8)), "w:color": color if color == "auto" else color.lstrip("#").upper(), "w:space": "0"}


def expected_table(table, rule):
    """Return an expected clone, or fail safely on ambiguous table structures."""
    expected = table.cloneNode(True)
    rows = children(expected, "w:tr")
    # Wrapped rows/cells (e.g. tracked revisions) need explicit handling first.
    if len(rows) != len([r for r in expected.getElementsByTagName("w:tr") if nearest_table(r) is expected]):
        raise ValueError("three-line tables with wrapped rows are unsupported")
    header_rows = rule.get("headerRows", 1)
    if not 1 <= header_rows < len(rows):
        raise ValueError("three-line tables need headerRows >= 1 and at least one body row")
    for i, row in enumerate(rows):
        cells = children(row, "w:tc")
        owned = [c for c in row.getElementsByTagName("w:tc") if nearest_table(c) is expected]
        if not cells or len(cells) != len(owned):
            raise ValueError("three-line tables with wrapped or missing cells are unsupported")
        if i == header_rows:
            for cell in cells:
                merge = child(child(cell, "w:tcPr"), "w:vMerge")
                if merge is not None and merge.getAttribute("w:val") != "restart":
                    raise ValueError("vertical merge crosses the three-line header boundary")
    color = rule.get("color", "000000")
    outer = border_values(rule.get("outerSizePt", 1.5), color)
    middle = border_values(rule.get("headerSizePt", 0.75), color)

    def borders(props, tag, values):
        target = ensure(props, tag)
        for old in list(target.childNodes):
            target.removeChild(old)
        for edge in EDGES + (["tl2br", "tr2bl"] if tag == "w:tcBorders" else []):
            node = target.ownerDocument.createElement("w:" + edge)
            for key, value in values.get(edge, {"w:val": "nil"}).items():
                node.setAttribute(key, value)
            target.appendChild(node)

    props = ensure(expected, "w:tblPr")
    configure_layout(expected, rule)
    borders(props, "w:tblBorders", {"top": outer, "bottom": outer})
    spacing = ensure(props, "w:tblCellSpacing")
    spacing.setAttribute("w:w", "0")
    spacing.setAttribute("w:type", "dxa")
    for i, row in enumerate(rows):
        row_props = ensure(row, "w:trPr")
        row_spacing = child(row_props, "w:tblCellSpacing")
        if row_spacing is not None:
            row_spacing.setAttribute("w:w", "0")
            row_spacing.setAttribute("w:type", "dxa")
        repeat = ensure(row_props, "w:tblHeader")
        repeat.setAttribute("w:val", "true" if i < header_rows and rule.get("repeatHeader", True) else "false")
        if "allowRowSplit" in rule:
            ensure(row_props, "w:cantSplit").setAttribute("w:val", "false" if rule["allowRowSplit"] else "true")
        for cell in children(row, "w:tc"):
            edges = {}
            if i == 0:
                edges["top"] = outer
            if i == header_rows:
                edges["top"] = middle
            if i == header_rows - 1:
                edges["bottom"] = middle
            if i == len(rows) - 1:
                edges["bottom"] = outer
            borders(ensure(cell, "w:tcPr"), "w:tcBorders", edges)
    configure_text(expected, rule)
    return expected


def owned_paragraphs(table):
    for paragraph in table.getElementsByTagName("w:p"):
        if nearest_table(paragraph) is not table:
            continue
        parent = paragraph.parentNode
        while parent is not table and parent.nodeName != "w:txbxContent":
            parent = parent.parentNode
        if parent is table:
            yield paragraph


def configure_text(table, rule):
    if not rule.get("paragraph") and not rule.get("run"):
        return
    # apply_profile imports this module; defer the shared setters until execution.
    from apply_profile import apply_fonts, apply_paragraph_format, ensure_direct, PPR_ORDER, RPR_ORDER
    for paragraph in owned_paragraphs(table):
        p_pr = ensure_direct(paragraph, "w:pPr")
        apply_paragraph_format(p_pr, rule.get("paragraph", {}), rule.get("run", {}).get("sizePt"))
        if rule.get("run"):
            apply_fonts(ensure_direct(p_pr, "w:rPr", PPR_ORDER), rule["run"])
            for run in paragraph_runs(paragraph):
                apply_fonts(ensure_direct(run, "w:rPr", RPR_ORDER), rule["run"])


def paragraph_runs(paragraph):
    for run in paragraph.getElementsByTagName("w:r"):
        parent = run.parentNode
        while parent is not None and parent.nodeName != "w:p":
            parent = parent.parentNode
        if parent is paragraph:
            yield run


def extra_checks(source, target, rule, parts):
    """Check only requested properties, with inherited text styles resolved."""
    from audit_docx import marginal_style_context, marginal_effective_props, run_props
    from marginal_format import RUN_KEYS, matches
    issues = []

    def attributes(node, names):
        return tuple(node.getAttribute(name) for name in names) if node is not None else None

    props, expected_props = child(source, "w:tblPr"), child(target, "w:tblPr")
    for key, tag, names in (("alignment", "w:jc", ("w:val",)), ("widthIn", "w:tblW", ("w:type", "w:w"))):
        if key in rule and attributes(child(props, tag), names) != attributes(child(expected_props, tag), names):
            issues.append(f"{key} differs")
    if "columnWidthsIn" in rule:
        def width_signature(table):
            props = child(table, "w:tblPr")
            grid = child(table, "w:tblGrid")
            return (attributes(child(props, "w:tblW"), ("w:type", "w:w")),
                    attributes(child(props, "w:tblLayout"), ("w:type",)),
                    [attributes(n, ("w:w",)) for n in children(grid, "w:gridCol")] if grid is not None else None,
                    [[attributes(child(child(c, "w:tcPr"), "w:tcW"), ("w:type", "w:w")) for c in children(r, "w:tc")] for r in children(table, "w:tr")])
        if width_signature(source) != width_signature(target):
            issues.append("columnWidthsIn differs")
    if "allowRowSplit" in rule:
        for i, row in enumerate(children(source, "w:tr")):
            node = child(child(row, "w:trPr"), "w:cantSplit")
            cannot_split = node is not None and node.getAttribute("w:val") not in {"false", "0", "off"}
            if cannot_split == rule["allowRowSplit"]:
                issues.append(f"row[{i}].allowRowSplit differs")
    context = marginal_style_context(parts)
    for i, paragraph in enumerate(owned_paragraphs(source)):
        runs = list(paragraph_runs(paragraph))
        actual_p, _, _ = marginal_effective_props(paragraph, runs[0] if runs else None, context)
        for key, value in rule.get("paragraph", {}).items():
            actual_key = key
            if key == "alignment" and value == "justify":
                value = "both"
            if key == "firstLineChars":
                actual_key = "firstLineIn"
                value *= rule.get("run", {}).get("sizePt", 12) / 72
            competing_chars = (key in {"firstLineIn", "hangingIn", "firstLineChars"} and value == 0
                               and any(actual_p.get(name) not in {None, 0} for name in ("firstLineChars", "hangingChars")))
            if competing_chars or not matches(actual_p.get(actual_key), value):
                issues.append(f"paragraph[{i}].{key} differs")
        actual_fonts = [marginal_effective_props(paragraph, run, context)[1] for run in runs]
        if not runs:
            actual_fonts = [run_props(child(child(paragraph, "w:pPr"), "w:rPr"))]
        for key, value in rule.get("run", {}).items():
            if key == "color":
                value = value.lstrip("#").upper()
            if any(not matches(actual.get(RUN_KEYS[key]), value) for actual in actual_fonts):
                issues.append(f"paragraph[{i}].run.{key} differs")
    return issues


def configure_layout(table, rule):
    props = ensure(table, "w:tblPr")
    if "alignment" in rule:
        ensure(props, "w:jc").setAttribute("w:val", rule["alignment"])
    widths = rule.get("columnWidthsIn")
    total = rule.get("widthIn", sum(widths) if widths else None)
    if total is not None:
        node = ensure(props, "w:tblW")
        node.setAttribute("w:type", "dxa")
        node.setAttribute("w:w", str(round(total * 1440)))
    if not widths:
        return
    twips = [round(value * 1440) for value in widths]
    ensure(props, "w:tblLayout").setAttribute("w:type", "fixed")
    grid = child(table, "w:tblGrid")
    if grid is None:
        grid = table.ownerDocument.createElement("w:tblGrid")
        table.insertBefore(grid, next((c for c in children(table) if c.nodeName != "w:tblPr"), None))
    for old in list(grid.childNodes):
        grid.removeChild(old)
    for value in twips:
        node = table.ownerDocument.createElement("w:gridCol")
        node.setAttribute("w:w", str(value))
        grid.appendChild(node)
    for row in children(table, "w:tr"):
        row_props = child(row, "w:trPr")
        if any(child(row_props, name) is not None for name in ("w:gridBefore", "w:gridAfter")):
            raise ValueError("column widths with omitted grid cells are unsupported")
        column = 0
        for cell in children(row, "w:tc"):
            cell_props = ensure(cell, "w:tcPr")
            if child(cell_props, "w:hMerge") is not None:
                raise ValueError("column widths with legacy hMerge are unsupported; use gridSpan")
            span_node = child(cell_props, "w:gridSpan")
            value = span_node.getAttribute("w:val") if span_node is not None else "1"
            if not value.isdigit() or int(value) < 1 or column + int(value) > len(twips):
                raise ValueError("columnWidthsIn does not match the table grid")
            span = int(value)
            node = ensure(cell_props, "w:tcW")
            node.setAttribute("w:type", "dxa")
            node.setAttribute("w:w", str(sum(twips[column:column + span])))
            column += span
        if column != len(twips):
            raise ValueError("columnWidthsIn does not match the table grid")


def nearest_table(node):
    parent = node.parentNode
    while parent is not None and parent.nodeName != "w:tbl":
        parent = parent.parentNode
    return parent


def apply_tables(document, rules, summary):
    if not rules:
        return
    tables = body_tables(document)
    edits = []
    for rule in rules:
        for index in rule["tableIndexes"]:
            if index >= len(tables):
                raise ValueError(f"table index {index} does not exist")
            source = tables[index]
            target = expected_table(source, rule)
            source.parentNode.replaceChild(target, source)
            tables[index] = target
            edits.append(index)
    if rules:
        summary["tablesConfigured"] = edits


def table_checks(document, rules, parts=None):
    if not rules:
        return []
    tables = body_tables(document)
    checks = []
    for rule_index, rule in enumerate(rules):
        issues = []
        for index in rule["tableIndexes"]:
            if index >= len(tables):
                issues.append(f"table {index}: missing")
                continue
            source = tables[index]
            try:
                target = expected_table(source, rule)
                if managed_signature(source) != managed_signature(target):
                    issues.append(f"table {index}: borders, spacing or repeating header differ")
                issues.extend(f"table {index}: {message}" for message in extra_checks(source, target, rule, parts or {}))
            except ValueError as exc:
                issues.append(f"table {index}: {exc}")
        checks.append({"name": f"tableRules[{rule_index}]", "ok": not issues, "expected": rule, "actual": issues})
    return checks


def managed_signature(table):
    """Compare managed properties, independent of XML whitespace/attribute order."""
    def borders(props, tag):
        root = child(props, tag)
        values = []
        for edge in EDGES + (["tl2br", "tr2bl"] if tag == "w:tcBorders" else []):
            node = child(root, "w:" + edge)
            if node is None:
                values.append(None)
                continue
            style = node.getAttribute("w:val")
            if style in {"nil", "none"}:
                values.append(("nil",))
            else:
                values.append((style, node.getAttribute("w:sz"), node.getAttribute("w:color").upper(), node.getAttribute("w:space") or "0"))
        return values

    def spacing(props):
        node = child(props, "w:tblCellSpacing")
        return (node.getAttribute("w:w"), node.getAttribute("w:type")) if node is not None else None

    props = child(table, "w:tblPr")
    values = [borders(props, "w:tblBorders"), spacing(props)]
    for row in children(table, "w:tr"):
        props = child(row, "w:trPr")
        header = child(props, "w:tblHeader")
        enabled = header is not None and header.getAttribute("w:val") not in {"0", "false", "off"}
        values.append((enabled, spacing(props), [borders(child(cell, "w:tcPr"), "w:tcBorders") for cell in children(row, "w:tc")]))
    return values
