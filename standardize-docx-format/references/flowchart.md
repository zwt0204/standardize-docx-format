# Editable native Word flowcharts

Use this workflow to turn a user's process description into a new flowchart. The skill interprets natural language; the Python command accepts structured JSON, not free-form prose. No formatting profile or external drawing service is needed.

## Interpret and generate

Identify each action, start/end, decision, branch label, merge, and return target. Preserve user terminology and branch direction. For example, “提交申请，审核通过则归档，不通过退回修改后重新提交” maps to [flowchart.example.json](flowchart.example.json). `title` names the drawing in Word's selection/metadata; put visible titles in the document separately if requested.

Run from the skill folder:

```bash
# New A4 DOCX (one-inch margins)
python scripts/standardize.py flowchart \
  --spec references/flowchart.example.json --output approval.docx

# Append to an existing DOCX, writing a new file
python scripts/standardize.py flowchart \
  --input report.docx --spec approval.json --output report.flowchart.docx

# Insert after a uniquely matched, exact top-level paragraph
python scripts/standardize.py flowchart \
  --input report.docx --spec approval.json --output report.flowchart.docx \
  --after-text '审批流程'
```

An insertion is additive: it does not replace the matched paragraph or other content. `--after-text` ignores paragraphs inside tables/text boxes and refuses ambiguous or absent matches and paragraphs carrying section breaks. Without it, the drawing goes before the final body section properties. Input and output must differ. Existing outputs require `--force`; the input and JSON spec remain protected even with that flag.

## JSON contract

The root accepts `title` (optional), `nodes`, `edges`, and optional `layout`. Unknown properties are errors.

- `nodes`: 1–60 objects, each with a unique ASCII `id` beginning with a letter, nonempty `text` (up to 200 characters), and optional `type`: `start`, `end`, `process` (default), or `decision`. They become rounded rectangles, rectangles, or diamonds with editable Word text.
- `edges`: up to 120 objects with existing node IDs in `from` and `to`, plus optional nonempty `label` (up to 40 characters; visual space usually permits only a short label such as “是/否”). Duplicate directed pairs are rejected. Cycles and self-loops are allowed. Nodes without edges are allowed.
- `row`, `column`: optional nonnegative integer grid positions on each node. Supply both on all nodes or omit both everywhere. Cells must be distinct. Automatic layout ranks the graph after excluding DFS back edges and centers branch rows. For manual layout, adjacent columns have a full node width plus column gap; larger grids may not fit a page.
- `layout`: optional `nodeWidthIn` (default 1.55, range 0.8–4), `nodeHeightIn` (0.7, range 0.4–2), `columnGapIn` (0.6, range 0.35–2), `rowGapIn` (0.55, range 0.35–2), `fontSizePt` (11, range 8–24), `font` (East Asian, default 宋体), `latinFont` (default Calibri). All node boxes share dimensions. Explicit `\n` splits node text into paragraphs.

Keep decision text short, since a diamond has a smaller text area. If the script reports overflow, shorten wording while preserving meaning, add deliberate line breaks, increase node dimensions, or split a large process into several diagrams. Do not silently shrink text to make a dense graph fit. Exterior lanes handle return paths and some long forward edges. Connectors crossing unrelated nodes and overlapping labels are rejected when detected; dense graphs can require explicit grid positions or splitting.

## Verification and boundaries

The output is an inline Word 2010+ grouped drawing (`wpg:wgp`). Nodes use `wps:cNvSpPr`, connectors use `wps:cNvCnPr` with `a:stCxn` / `a:endCxn` references, and labels use editable text boxes. It does not create SmartArt or raster images, and it does not edit existing diagrams.

Only `word/document.xml` is rewritten in an existing package. Other parts, relationships, images, and compatibility branches are preserved. The command checks section content dimensions and refuses oversized canvases and multi-column sections rather than overflowing the page. It estimates text fit; actual fonts, inherited formatting, pagination, connector routing, and other drawing applications still require rendering review. There is no Word 2007/VML fallback.

Audit and optionally render the generated DOCX:

```bash
python scripts/standardize.py audit --input approval.docx --output approval.audit.json
python scripts/standardize.py visual --input approval.docx \
  --output approval.visual.json --render-dir approval-rendered
```

The command reports node/edge counts, canvas size, and selected grid under `summary`. `ok` means generation completed; `visualConfirmation` stays `pending`. Check all node text, arrow direction, decision labels, return paths, node/line overlap, and page fit. Existing thesis-oriented visual heuristics can warn about missing chapters or figure captions in a standalone diagram; assess these against the user's request. LibreOffice rendering is useful evidence but cannot prove Microsoft Word opens without a repair prompt or keeps connectors attached after dragging; only claim those behaviors after testing in Word.

Deliver the output DOCX, reusable JSON spec, and the actual verification result, including any unavailable Word check.
