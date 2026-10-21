"""Generate editable Word 2010+ grouped drawings from a small flowchart spec.

Natural-language interpretation belongs to the skill; this module is deterministic
and uses only the standard library. Existing package parts are preserved verbatim
except for document.xml, which is edited using a compatibility-preserving DOM.
"""

from __future__ import annotations

import math
import re
import zipfile
from collections import deque
from pathlib import Path
from typing import Any
from xml.dom import Node, minidom

from apply_profile import direct_child, element_children, validate_package_parts, write_package

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
EMU = 914400
SHAPES = {"start": "roundRect", "end": "roundRect", "process": "rect", "decision": "diamond"}
PORT_INDEX = {"top": 0, "left": 1, "bottom": 2, "right": 3}


def keys(value: Any, allowed: set[str], where: str) -> None:
    if not isinstance(value, dict):
        raise ValueError(f"{where} must be an object")
    unknown = value.keys() - allowed
    if unknown:
        raise ValueError(f"Unknown {where} keys: {', '.join(sorted(unknown))}")


def text_value(value: Any, where: str, limit: int = 200) -> str:
    if not isinstance(value, str) or not value.strip() or len(value) > limit:
        raise ValueError(f"{where} must be nonempty text of at most {limit} characters")
    if any(not (c in "\t\n\r" or 0x20 <= ord(c) <= 0xD7FF or
                   0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF) for c in value):
        raise ValueError(f"{where} contains characters invalid in XML")
    return value


def number(value: Any, where: str, low: float, high: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or not low <= value <= high:
        raise ValueError(f"{where} must be a finite number between {low} and {high}")
    return float(value)


def validate_spec(spec: dict[str, Any]) -> None:
    keys(spec, {"title", "nodes", "edges", "layout"}, "flowchart")
    text_value(spec.get("title", "Flowchart"), "title")
    nodes, edges = spec.get("nodes"), spec.get("edges")
    if not isinstance(nodes, list) or not 1 <= len(nodes) <= 60:
        raise ValueError("nodes must contain 1 to 60 nodes")
    if not isinstance(edges, list) or len(edges) > 120:
        raise ValueError("edges must be a list of at most 120 edges")
    ids, cells = set(), set()
    positioned = []
    for node in nodes:
        keys(node, {"id", "text", "type", "row", "column"}, "node")
        nid = text_value(node.get("id"), "node.id", 80)
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]*", nid) or nid in ids:
            raise ValueError(f"Invalid or duplicate node id: {nid}")
        ids.add(nid)
        text_value(node.get("text"), f"node {nid} text")
        if node.get("type", "process") not in SHAPES:
            raise ValueError(f"Unsupported node type: {node.get('type')}")
        has_position = "row" in node or "column" in node
        positioned.append(has_position)
        if has_position:
            for key in ("row", "column"):
                value = node.get(key)
                if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value <= 59:
                    raise ValueError(f"node {nid} {key} must be an integer from 0 to 59")
            cell = (node["row"], node["column"])
            if cell in cells:
                raise ValueError(f"Overlapping grid position: {cell}")
            cells.add(cell)
    if any(positioned) and not all(positioned):
        raise ValueError("Specify row and column for every node, or omit them for every node")
    pairs = set()
    for edge in edges:
        keys(edge, {"from", "to", "label"}, "edge")
        for key in ("from", "to"):
            if not isinstance(edge.get(key), str) or edge[key] not in ids:
                raise ValueError(f"Unknown edge {key}: {edge.get(key)}")
        pair = (edge["from"], edge["to"])
        if pair in pairs:
            raise ValueError(f"Duplicate edge: {pair}; combine its labels")
        pairs.add(pair)
        if "label" in edge:
            text_value(edge["label"], "edge.label", 40)
    layout = spec.get("layout", {})
    keys(layout, {"nodeWidthIn", "nodeHeightIn", "columnGapIn", "rowGapIn", "fontSizePt", "font", "latinFont"}, "layout")
    for key, default, lo, hi in (
        ("nodeWidthIn", 1.55, .8, 4), ("nodeHeightIn", .7, .4, 2),
        ("columnGapIn", .6, .35, 2), ("rowGapIn", .55, .35, 2),
        ("fontSizePt", 11, 8, 24),
    ):
        number(layout.get(key, default), key, lo, hi)
    for key in ("font", "latinFont"):
        if key in layout:
            text_value(layout[key], key, 80)


def grid_positions(spec: dict[str, Any]) -> dict[str, tuple[int, int]]:
    """Longest-path ranks after DFS back-edge removal; cycles become return lanes."""
    nodes = spec["nodes"]
    if "row" in nodes[0]:
        return {n["id"]: (n["row"], n["column"]) for n in nodes}
    adjacency = {n["id"]: [] for n in nodes}
    for edge in spec["edges"]:
        adjacency[edge["from"]].append(edge["to"])
    state, forward = {}, {nid: [] for nid in adjacency}

    def visit(nid):
        state[nid] = 1
        for target in adjacency[nid]:
            if state.get(target) == 1:
                continue
            forward[nid].append(target)
            if not state.get(target):
                visit(target)
        state[nid] = 2

    for nid in adjacency:
        if not state.get(nid):
            visit(nid)
    degree, rank = dict.fromkeys(adjacency, 0), dict.fromkeys(adjacency, 0)
    for targets in forward.values():
        for target in targets:
            degree[target] += 1
    queue = deque(nid for nid in adjacency if not degree[nid])
    while queue:
        nid = queue.popleft()
        for target in forward[nid]:
            rank[target] = max(rank[target], rank[nid] + 1)
            degree[target] -= 1
            if not degree[target]:
                queue.append(target)
    rows = {}
    for nid in adjacency:
        rows.setdefault(rank[nid], []).append(nid)
    max_columns = max(map(len, rows.values()))
    # Half-grid columns center single nodes above/below two-way branches.
    return {nid: (row, 2 * col + max_columns - len(group))
            for row, group in rows.items() for col, nid in enumerate(group)}


def element(document, parent, tag, attrs=None):
    node = document.createElement(tag)
    for key, value in (attrs or {}).items():
        node.setAttribute(key, str(value))
    parent.appendChild(node)
    return node


def textbox(document, shape, text, layout):
    box = element(document, shape, "wps:txbx")
    content = element(document, box, "w:txbxContent")
    for line in text.split("\n"):
        p = element(document, content, "w:p")
        props = element(document, p, "w:pPr")
        element(document, props, "w:spacing", {"w:before": 0, "w:after": 0, "w:line": 240, "w:lineRule": "auto"})
        element(document, props, "w:ind", {"w:left": 0, "w:right": 0, "w:firstLine": 0})
        element(document, props, "w:jc", {"w:val": "center"})
        run = element(document, p, "w:r")
        rp = element(document, run, "w:rPr")
        element(document, rp, "w:rFonts", {"w:ascii": layout.get("latinFont", "Calibri"),
                "w:hAnsi": layout.get("latinFont", "Calibri"), "w:eastAsia": layout.get("font", "宋体")})
        element(document, rp, "w:color", {"w:val": "172B4D"})
        element(document, rp, "w:sz", {"w:val": round(layout.get("fontSizePt", 11) * 2)})
        t = element(document, run, "w:t", {"xml:space": "preserve"})
        t.appendChild(document.createTextNode(line))
    # Explicitly disable automatic text resizing; oversized labels are rejected.
    body = element(document, shape, "wps:bodyPr", {"rot": 0, "vert": "horz", "wrap": "square",
                   "lIns": 18000, "tIns": 9000, "rIns": 18000, "bIns": 9000, "anchor": "ctr"})
    element(document, body, "a:noAutofit")


def shape_base(document, canvas, sid, name, box, *, connector=False):
    shape = element(document, canvas, "wps:wsp")
    element(document, shape, "wps:cNvPr", {"id": sid, "name": name})
    nv = element(document, shape, "wps:cNvCnPr" if connector else "wps:cNvSpPr")
    props = element(document, shape, "wps:spPr")
    transform = element(document, props, "a:xfrm")
    x, y, width, height = box
    element(document, transform, "a:off", {"x": x, "y": y})
    element(document, transform, "a:ext", {"cx": width, "cy": height})
    return shape, nv, props


def solid(document, parent, color):
    fill = element(document, parent, "a:solidFill")
    element(document, fill, "a:srgbClr", {"val": color})


def line_style(document, props, *, arrow=False, hidden=False):
    line = element(document, props, "a:ln", {"w": 15240})
    if hidden:
        element(document, line, "a:noFill")
    else:
        solid(document, line, "42526E")
    if arrow:
        element(document, line, "a:tailEnd", {"type": "triangle", "w": "sm", "len": "sm"})


def text_fits(text, width, height, font, *, decision=False):
    """Conservative estimate, not a replacement for Word text layout."""
    if decision:
        width, height = width * .5, height * .5
    usable = (width - .08) * 72 / font
    lines = 0
    for line in text.split("\n"):
        units = sum(1 if ord(c) > 255 else .58 for c in line)
        lines += max(1, math.ceil(units / max(.01, usable)))
    return lines * font * 1.2 <= (height - .04) * 72


def crosses_box(points, box):
    x, y, w, h = box
    for (ax, ay), (bx, by) in zip(points, points[1:]):
        if ax == bx and x < ax < x + w and max(min(ay, by), y) < min(max(ay, by), y + h):
            return True
        if ay == by and y < ay < y + h and max(min(ax, bx), x) < min(max(ax, bx), x + w):
            return True
    return False


def build_canvas(document, spec, next_id, max_width, max_height):
    layout = spec.get("layout", {})
    nw, nh = layout.get("nodeWidthIn", 1.55), layout.get("nodeHeightIn", .7)
    cg, rg = layout.get("columnGapIn", .6), layout.get("rowGapIn", .55)
    font = layout.get("fontSizePt", 11)
    grid = grid_positions(spec)
    auto = "row" not in spec["nodes"][0]
    boxes, ids = {}, {}
    pad = .12
    for offset, node in enumerate(spec["nodes"]):
        row, column = grid[node["id"]]
        x = pad + column * (nw + cg) / (2 if auto else 1)
        y = pad + row * (nh + rg)
        if not text_fits(node["text"], nw, nh, font, decision=node.get("type") == "decision"):
            raise ValueError(f"Text may overflow node {node['id']}; shorten it, add line breaks, or increase node dimensions")
        boxes[node["id"]] = tuple(round(v * EMU) for v in (x, y, nw, nh))
        ids[node["id"]] = next_id + 1 + offset
    right = max(x + w for x, y, w, h in boxes.values())
    bottom = max(y + h for x, y, w, h in boxes.values())
    lane = 0
    routes = []

    def port(box, side):
        x, y, w, h = box
        return {"top": (x + w // 2, y), "bottom": (x + w // 2, y + h),
                "left": (x, y + h // 2), "right": (x + w, y + h // 2)}[side]

    for edge in spec["edges"]:
        source, target = boxes[edge["from"]], boxes[edge["to"]]
        if target[1] > source[1]:
            ss, ts = ("bottom" if target[0] == source[0] else
                      "right" if target[0] > source[0] else "left"), "top"
            start, end = port(source, ss), port(target, ts)
            if ss == "bottom":
                mid = (start[1] + end[1]) // 2
                points = [start, (start[0], mid), (end[0], mid), end]
            else:
                points = [start, (end[0], start[1]), end]
        elif target[1] == source[1] and target[0] != source[0]:
            ss, ts = ("right", "left") if target[0] > source[0] else ("left", "right")
            start, end = port(source, ss), port(target, ts)
            points = [start, end]
        else:
            # Separate lanes outside all nodes make return arrows and self-loops visible.
            lane += 1
            ss, ts = "right", "right"
            start, end = port(source, ss), port(target, ts)
            lx = right + round((.28 + lane * .18) * EMU)
            if edge["from"] == edge["to"]:
                ts = "top"
                end = port(target, ts)
                top = target[1] - round(.09 * EMU)
                points = [start, (lx, start[1]), (lx, top), (end[0], top), end]
            else:
                points = [start, (lx, start[1]), (lx, end[1]), end]
        if any(crosses_box(points, box) for nid, box in boxes.items()
               if nid not in {edge["from"], edge["to"]}):
            # A long edge may skip over occupied ranks. Try an exterior lane.
            lane += 1
            ss, ts = "right", "right"
            start, end = port(source, ss), port(target, ts)
            lx = right + round((.28 + lane * .18) * EMU)
            points = [start, (lx, start[1]), (lx, end[1]), end]
            if any(crosses_box(points, box) for nid, box in boxes.items()
                   if nid not in {edge["from"], edge["to"]}):
                raise ValueError("A connector crosses an unrelated node; adjust explicit row/column positions or split the flowchart")
        routes.append((edge, ss, ts, points))
    width = max([right] + [x for _, _, _, pts in routes for x, y in pts]) + round(pad * EMU)
    height = bottom + round(pad * EMU)
    if width > max_width * EMU or height > (max_height - .15) * EMU:
        raise ValueError(f"Flowchart ({width / EMU:.2f} x {height / EMU:.2f}in) exceeds section content box "
                         f"({max_width:.2f} x {max_height:.2f}in); reduce grid gaps/node dimensions or split the flowchart")

    p = document.createElement("w:p")
    # Scope extension namespaces to the inserted subtree; preserve root declarations.
    for prefix, uri in (("w", W), ("wp", WP), ("a", A), ("wps", WPS), ("wpg", WPG)):
        p.setAttribute("xmlns:" + prefix, uri)
    pp = element(document, p, "w:pPr")
    element(document, pp, "w:keepLines")
    element(document, pp, "w:spacing", {"w:before": 0, "w:after": 0, "w:line": 240, "w:lineRule": "auto"})
    element(document, pp, "w:ind", {"w:left": 0, "w:right": 0, "w:firstLine": 0})
    element(document, pp, "w:jc", {"w:val": "center"})
    drawing = element(document, element(document, p, "w:r"), "w:drawing")
    inline = element(document, drawing, "wp:inline", {"distT": 0, "distB": 0, "distL": 0, "distR": 0})
    element(document, inline, "wp:extent", {"cx": width, "cy": height})
    element(document, inline, "wp:docPr", {"id": next_id, "name": spec.get("title", "Flowchart"), "descr": "Editable flowchart"})
    graphic = element(document, inline, "a:graphic")
    data = element(document, graphic, "a:graphicData", {"uri": WPG})
    canvas = element(document, data, "wpg:wgp")
    element(document, canvas, "wpg:cNvGrpSpPr")
    group_props = element(document, canvas, "wpg:grpSpPr")
    transform = element(document, group_props, "a:xfrm")
    element(document, transform, "a:off", {"x": 0, "y": 0})
    element(document, transform, "a:ext", {"cx": width, "cy": height})
    element(document, transform, "a:chOff", {"x": 0, "y": 0})
    element(document, transform, "a:chExt", {"cx": width, "cy": height})
    # Draw connectors behind the nodes. Their references may point forward.
    sid = next_id + 1 + len(ids)
    labels = []
    for edge, ss, ts, points in routes:
        x, y = min(p[0] for p in points), min(p[1] for p in points)
        w, h = max(1, max(p[0] for p in points) - x), max(1, max(p[1] for p in points) - y)
        shape, nv, props = shape_base(document, canvas, sid, f"{edge['from']} -> {edge['to']}", (x, y, w, h), connector=True)
        sid += 1
        element(document, nv, "a:stCxn", {"id": ids[edge["from"]], "idx": PORT_INDEX[ss]})
        element(document, nv, "a:endCxn", {"id": ids[edge["to"]], "idx": PORT_INDEX[ts]})
        geometry = element(document, props, "a:custGeom")
        for tag in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
            element(document, geometry, tag)
        element(document, geometry, "a:rect", {"l": "l", "t": "t", "r": "r", "b": "b"})
        path = element(document, element(document, geometry, "a:pathLst"), "a:path", {"w": w, "h": h, "fill": "none"})
        for i, (px, py) in enumerate(points):
            step = element(document, path, "a:moveTo" if i == 0 else "a:lnTo")
            element(document, step, "a:pt", {"x": px - x, "y": py - y})
        element(document, props, "a:noFill")
        line_style(document, props, arrow=True)
        element(document, shape, "wps:bodyPr")
        if edge.get("label"):
            # A label near the outgoing segment avoids the merge point.
            px, py = points[0]
            if ss == "bottom":
                lx, ly = px + round(.04 * EMU), py + round(.02 * EMU)
            elif ss == "left":
                lx, ly = px - round(.76 * EMU), py - round(.25 * EMU)
            else:
                lx, ly = px + round(.04 * EMU), py - round(.25 * EMU)
            label_width = min(.72, (width - lx) / EMU - .02)
            if lx < 0 or ly < 0:
                raise ValueError("Edge label would extend outside the canvas; adjust node grid")
            if label_width < .2 or not text_fits(edge["label"], label_width, .24, font):
                raise ValueError(f"Edge label {edge['label']!r} is too long; shorten it")
            labels.append((edge["label"], (lx, ly, round(label_width * EMU), round(.24 * EMU))))

    def overlaps(one, two):
        x, y, w, h = one
        ox, oy, ow, oh = two
        return max(x, ox) < min(x + w, ox + ow) and max(y, oy) < min(y + h, oy + oh)

    for i, (label, box) in enumerate(labels):
        if any(overlaps(box, other) for other in boxes.values()) or any(overlaps(box, other) for _, other in labels[:i]):
            raise ValueError(f"Edge label {label!r} overlaps a node or another label; adjust the grid or shorten labels")

    for node in spec["nodes"]:
        shape, nv, props = shape_base(document, canvas, ids[node["id"]], node["id"], boxes[node["id"]])
        geometry = element(document, props, "a:prstGeom", {"prst": SHAPES[node.get("type", "process")]})
        element(document, geometry, "a:avLst")
        solid(document, props, "FFF4CC" if node.get("type") == "decision" else "EAF2FF")
        line_style(document, props)
        textbox(document, shape, node["text"], layout)
    for label, box in labels:
        shape, nv, props = shape_base(document, canvas, sid, "Branch label", box)
        sid += 1
        geometry = element(document, props, "a:prstGeom", {"prst": "rect"})
        element(document, geometry, "a:avLst")
        solid(document, props, "FFFFFF")
        line_style(document, props, hidden=True)
        textbox(document, shape, label, layout)
    return p, {"nodes": len(ids), "edges": len(routes), "widthIn": round(width / EMU, 3),
               "heightIn": round(height / EMU, 3), "grid": {k: list(v) for k, v in grid.items()}}


def blank_package():
    return {
        "[Content_Types].xml": b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        "_rels/.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/document.xml": f'<w:document xmlns:w="{W}"><w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>'.encode(),
    }


def section_box(body, reference):
    children = element_children(body)
    start = children.index(reference) if reference else len(children)
    section = None
    for child in children[start:]:
        section = child if child.nodeName == "w:sectPr" else direct_child(direct_child(child, "w:pPr"), "w:sectPr")
        if section is not None:
            break
    if section is None:
        raise ValueError("Cannot determine insertion section geometry")
    size, margins = direct_child(section, "w:pgSz"), direct_child(section, "w:pgMar")
    cols = direct_child(section, "w:cols")
    if cols is not None and (int(cols.getAttribute("w:num") or "1") > 1 or
                             len(cols.getElementsByTagName("w:col")) > 1):
        raise ValueError("Multi-column insertion sections are not supported")
    if size is None or margins is None:
        raise ValueError("Insertion section must specify page size and margins")
    try:
        width = int(size.getAttribute("w:w")) - sum(int(margins.getAttribute("w:" + k) or "0") for k in ("left", "right", "gutter"))
        height = int(size.getAttribute("w:h")) - sum(int(margins.getAttribute("w:" + k)) for k in ("top", "bottom"))
    except ValueError as exc:
        raise ValueError("Invalid insertion section dimensions") from exc
    if min(width, height) <= 0:
        raise ValueError("Insertion section has no usable page area")
    return width / 1440, height / 1440


def add_flowchart(spec: dict[str, Any], output: Path, *, source: Path | None = None,
                  after_text: str | None = None, force: bool = False) -> dict[str, Any]:
    validate_spec(spec)
    if source is not None and (source.resolve() == output.resolve() or
            (output.exists() and source.samefile(output))):
        raise ValueError("Refusing in-place write: input and output refer to the same file")
    if after_text is not None and source is None:
        raise ValueError("--after-text requires --input")
    infos = []
    if source is None:
        parts = blank_package()
    else:
        with zipfile.ZipFile(source) as archive:
            infos = archive.infolist()
            parts = {info.filename: archive.read(info) for info in infos}
    validate_package_parts(parts)
    document = minidom.parseString(parts["word/document.xml"])
    bodies = document.getElementsByTagName("w:body")
    if len(bodies) != 1:
        raise ValueError("Expected one w:body in word/document.xml")
    body = bodies[0]
    section = direct_child(body, "w:sectPr")
    reference = section
    if after_text is not None:
        matches = []
        for child in element_children(body):
            if child.nodeName != "w:p" or child.getElementsByTagName("w:txbxContent"):
                continue
            text = "".join(t.firstChild.data if t.firstChild else "" for t in child.getElementsByTagName("w:t"))
            if text == after_text:
                matches.append(child)
        if len(matches) != 1:
            raise ValueError(f"--after-text must match exactly one top-level paragraph; matched {len(matches)}")
        anchor = matches[0]
        if direct_child(direct_child(anchor, "w:pPr"), "w:sectPr") is not None:
            raise ValueError("Cannot insert after a paragraph carrying a section break; select a paragraph within the target section")
        reference = anchor.nextSibling
        while reference is not None and reference.nodeType != Node.ELEMENT_NODE:
            reference = reference.nextSibling
    max_width, max_height = section_box(body, reference)
    used = []
    for name, data in parts.items():
        if name.endswith(".xml"):
            dom = minidom.parseString(data)
            for node in dom.getElementsByTagName("*"):
                if node.localName in {"docPr", "cNvPr"} and node.getAttribute("id").isdigit():
                    used.append(int(node.getAttribute("id")))
    paragraph, summary = build_canvas(document, spec, max(used, default=0) + 1, max_width, max_height)
    body.insertBefore(paragraph, reference)
    parts["word/document.xml"] = document.toxml(encoding="utf-8")
    validate_package_parts(parts)
    write_package(infos, parts, output, force)
    return {"ok": True, "output": str(output.resolve()), "summary": summary,
            "visualConfirmation": "pending", "compatibility": "Editable grouped drawing requires Word 2010 or newer"}
