# standardize-docx-format

[简体中文](README.md) | [English](README.en.md)

A Codex Skill for standardizing Microsoft Word `.docx` files. It targets theses, dissertations, course papers, and technical reports: complex layout, official-template filling, and machine-checkable acceptance.

It can follow an official Word template, or compile natural-language / PDF / web / screenshot rules into a structured JSON profile and apply that profile deterministically to an existing DOCX.

The current implementation is **v2.1**: the same deterministic OOXML engine, plus Profile JSON Schema, a selectable Repair Plan, deeper template analysis, structural validation, and Visual QA.

```text
spec / template → Requirement IR / Document Spec → Profile
      → Audit / Semantic AST → Diff → Repair Plan
      → confirmed Apply → Validate → Visual QA
```

The LLM interprets the specification. OOXML edits stay deterministic. Diff and Repair Plan are shown first; documents are not mutated silently.

> This project is a Codex Skill, not a Word add-in and not a GUI typesetter. Core scripts use the Python standard library only. They can audit a DOCX package, build a semantic model, diff, plan repairs, apply a profile, validate structure, and run visual heuristics. Ambiguous requirement extraction and template-region mapping stay in the Codex workflow. Final pagination is still judged in Word or LibreOffice.

The runtime contract for Codex is [`standardize-docx-format/SKILL.md`](standardize-docx-format/SKILL.md).

## Capabilities

- Keep the source DOCX; write a new `.standardized.docx`; refuse in-place overwrite.
- Preserve images, tables, bookmarks, fields, relationships, headers/footers, and unmodeled OOXML.
- Page sizes A3/A4/A5/Letter, orientation, and per-section margins.
- Different section rules for cover, abstract, TOC, body, references, acknowledgements, appendix.
- Separate East Asian and Latin font slots (`w:eastAsia` vs `w:ascii` / `w:hAnsi` / `w:cs`).
- Font size, bold/italic/color, spacing, first-line/hanging indent, justification, pagination controls.
- Roman vs decimal page numbers, with body restart at 1.
- Independent first / default / even headers and footers. Official templates can use `action: inherit` so existing header/footer refs are not deleted and cover coordinates are not rebuilt.
- Chinese multilevel numbering: `一、`, `（一）`, `1.`, `（1）`, `①`.
- TOC, PAGE, NUMPAGES, SEQ, REF, PAGEREF, STYLEREF, CITATION, BIBLIOGRAPHY as Word fields.
- Bookmarks, captions, figure/table numbering, cross-reference infrastructure; chapter-number strategy via `captionNumbering`.
- Exact placeholder replacement across runs, headers, footers, and text boxes.
- Semantic style mapping by text, regex, current style, section, paragraph index, table/text-box location. TOC styles can be excluded with `styleNameNotRegex`.
- Strip template instructions and handwritten heading numbers; unify body formatting.
- Validate structure, required chapters, fields, bookmarks, page numbers, headers/footers, placeholders, and abstract word counts.
- Lift a thesis into a Document Semantic AST (cover, abstracts, TOC, chapters, captions, references).
- Profile Diff and Repair Plan: explain expected vs actual, apply only after confirmation; `--only-auto` or `--only-ids`.
- Visual QA in three layers: OOXML heuristics, pagination estimates, optional `docx-cli` / LibreOffice render.
- Compile a draft profile from spec text/PDF; reverse a draft from an official template (numbering, TOC, placeholders, bounded `paragraphRules`).
- Validate profiles against `references/profile.schema.json`.
- Backward-compatible v1 profiles plus v2 Profile / Document Spec.

## When to use it

### 1. Official DOCX template

Use the school/org `.docx` as authority when:

- The cover uses exact underlines, text boxes, tabs, or floating objects.
- Front matter uses Roman numerals and the body restarts at Arabic 1.
- Odd/even headers differ.
- TOC, captions, and cross-references must remain fields.
- The template has declaration pages, signature areas, content controls, or unusual sections.

Fill placeholders and map thesis content onto semantic regions. Do not rebuild encoded cover layout.

### 2. Spec only (no template)

Typical inputs: published text rules, a PDF guide, a web page, screenshots, or a verbal description of fonts, spacing, page numbers, and heading rules.

Codex translates measurable rules into a JSON profile. Ambiguous or conflicting items go here:

```json
{
  "requirements": {
    "unresolved": ["cover fill-in line length is unspecified"],
    "conflicts": []
  }
}
```

Apply refuses by default while `unresolved` or `conflicts` is non-empty.

### 3. Reusable profile

Generate a profile once per school/journal, then audit / apply / validate many papers, keeping only cover variables and a few locators per document.

## Repository layout

```text
.
├── README.md
├── README.en.md
├── .github/workflows/ci.yml
├── dist/
│   └── standardize-docx-format-v1.0.0.zip   # historical zip name; trust current source
└── standardize-docx-format/                 # the Skill directory to install
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    ├── scripts/
    │   ├── standardize.py          # unified CLI / pipeline
    │   ├── profile_schema.py
    │   ├── plan_apply.py
    │   ├── audit_docx.py
    │   ├── document_model.py
    │   ├── compile_requirements.py
    │   ├── analyze_template.py
    │   ├── diff_profile.py
    │   ├── repair_plan.py
    │   ├── apply_profile.py
    │   ├── validate_docx.py
    │   ├── visual_qa.py
    │   ├── layout_estimate.py
    │   ├── visual_repair.py
    │   ├── report.py
    │   ├── test_paragraph_match_safety.py
    │   └── test_v2_pipeline.py
    └── references/
        ├── v2-architecture.md
        ├── profile.schema.json
        ├── profile-schema.md
        ├── spec-driven-workflow.md
        ├── ooxml-advanced.md
        ├── thesis-cn.example.json
        └── guangzhou-nanfang-thesis.example.json
```

Root README files and `dist/` are not part of the runtime Skill. GitHub Actions compiles scripts, validates example profiles, and runs regression tests.

## Requirements

Required:

- Python **3.10+**. CI currently runs 3.11 and 3.12. Do not use 3.9 or older.
- Local read/write access to `.docx` files.
- Codex (or a compatible Skills host) if you use it as a Skill. The scripts also run standalone.

Optional:

- Microsoft Word: field refresh and final visual sign-off (source of truth).
- LibreOffice / `soffice`: headless DOCX → PDF.
- `pdftoppm` (poppler): PDF → PNG for Visual QA.
- `docx-cli` (`docx`): locators, diffs, and render. Not required by core scripts. Prefer `docx render` when present.

No `python-docx`, `lxml`, or other third-party Python packages.

## Install

### Option 1: ZIP

Download a zip from Releases or `dist/`. The file in this repo may still be named `standardize-docx-format-v1.0.0.zip`. **Treat `standardize-docx-format/` in git as the current version**, not the zip filename. Copy the inner `standardize-docx-format` folder into the Codex Skills directory.

Windows:

```text
C:\Users\<you>\.codex\skills\standardize-docx-format
```

macOS/Linux:

```text
~/.codex/skills/standardize-docx-format
```

With `CODEX_HOME`:

```text
$CODEX_HOME/skills/standardize-docx-format
```

### Option 2: clone

```bash
git clone https://github.com/zwt0204/standardize-docx-format.git
```

Windows PowerShell:

```powershell
$skillsRoot = if ($env:CODEX_HOME) {
  Join-Path $env:CODEX_HOME "skills"
} else {
  Join-Path $env:USERPROFILE ".codex\skills"
}

Copy-Item `
  -LiteralPath ".\standardize-docx-format\standardize-docx-format" `
  -Destination $skillsRoot `
  -Recurse
```

macOS/Linux:

```bash
skills_root="${CODEX_HOME:-$HOME/.codex}/skills"
mkdir -p "$skills_root"
cp -R ./standardize-docx-format/standardize-docx-format "$skills_root/"
```

Restart the Codex session (or rescan Skills) after install.

## Use inside Codex

```text
Use $standardize-docx-format to format thesis.docx from university-template.docx. Keep the original and write a new file.
```

```text
Use $standardize-docx-format to compile a profile from this formatting spec, list unresolved items, and apply only after confirmation.
```

```text
Use $standardize-docx-format to audit thesis.docx (sections, fonts, page numbers, TOC, captions, references) without modifying the file.
```

```text
Use $standardize-docx-format to standardize this batch of DOCX files with profile.json and emit a validation report for each.
```

## CLI quick start

Always write to a new output path. PowerShell examples below; on Linux/macOS replace backticks with `\` and fix paths.

### 1. Audit

```powershell
python .\standardize-docx-format\scripts\standardize.py audit `
  --input "D:\docs\thesis.docx" `
  --output ".\before-audit.json"
```

The audit covers ZIP integrity, paragraph/run/table/image counts, style definitions vs use, direct formatting, sections and margins, page-number format/restart, header/footer refs, fields, bookmarks, drawings, text boxes, content controls, and CJK/Latin counts.

### 2. Schema, diff, repair plan

```powershell
python .\standardize-docx-format\scripts\standardize.py schema --input ".\profile.json"
python .\standardize-docx-format\scripts\standardize.py diff `
  --input "D:\docs\thesis.docx" --profile ".\profile.json" --text
python .\standardize-docx-format\scripts\standardize.py plan `
  --input "D:\docs\thesis.docx" --profile ".\profile.json" --text
```

Do not `apply` until the plan is confirmed.

- `--only-auto` runs auto-applyable steps only.
- `map_structure` (missing abstract/references, etc.) and `pageBreakBefore` are **not** auto; pass their ids in `--only-ids`.
- A `--plan` apply also requires `--yes`.

### 3. Apply

Full profile:

```powershell
python .\standardize-docx-format\scripts\apply_profile.py `
  --input "D:\docs\thesis.docx" `
  --output "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json"
```

Confirmed auto steps only:

```powershell
python .\standardize-docx-format\scripts\standardize.py apply `
  --input "D:\docs\thesis.docx" `
  --output "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --plan ".\qa-output\repair-plan.json" `
  --only-auto --yes --force
```

Safety: refuse same input/output path; refuse overwrite unless `--force`; refuse unresolved/conflicts unless `--allow-unresolved`; re-check ZIP/XML after write.

### 4. Structural validate

```powershell
python .\standardize-docx-format\scripts\validate_docx.py `
  --input "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --output ".\validation.json"
```

Non-zero exit on failure. Structural pass is not visual pass.

### 5. Visual QA and real render

Heuristics and pagination estimates only (no Word):

```powershell
python .\standardize-docx-format\scripts\standardize.py visual `
  --input "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --output ".\qa-output\visual.json" `
  --html ".\qa-output\visual-report.html"
```

PDF/PNG requires `--render-dir`. `standardize.py visual` has **no** `--render` flag (`pipeline` does):

```powershell
python .\standardize-docx-format\scripts\standardize.py visual `
  --input "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --render-dir ".\qa-output\rendered" `
  --output ".\qa-output\visual.json" `
  --html ".\qa-output\visual-report.html"
```

Render order:

1. `docx render … --out <render-dir>/rendered` if `docx` exists
2. Else LibreOffice `--headless --convert-to pdf`, then `pdftoppm -png` when available
3. Else `render.skipped=true`; heuristics still run

The machine only flags tiny PNGs vs the median (possible blank page) and tiny PDFs (possible failed convert). Header collisions, table overflow, cover shift, TOC page numbers, CJK numbering, and STYLEREF need a human (or an agent reading the PNGs) plus Word field refresh.

```powershell
docx render output.standardized.docx --out rendered-pages
```

Inspect at least: cover, TOC, first body page, a figure/table page, last page. In Word: `Ctrl+A` then `F9` (update entire TOC if asked). A LibreOffice pass does not prove Word TOC/STYLEREF results.

### 6. Compiler pipeline

```powershell
python .\standardize-docx-format\scripts\standardize.py compile `
  --input ".\spec.txt" `
  --output ".\requirement-ir.json" `
  --profile-out ".\profile.json"

python .\standardize-docx-format\scripts\standardize.py analyze-template `
  --input ".\university-template.docx" `
  --profile-out ".\template-profile.json"

python .\standardize-docx-format\scripts\standardize.py pipeline `
  --input "D:\docs\thesis.docx" `
  --profile ".\profile.json" `
  --work-dir ".\qa-output"

python .\standardize-docx-format\scripts\standardize.py pipeline `
  --input "D:\docs\thesis.docx" `
  --profile ".\profile.json" `
  --output "D:\docs\thesis.standardized.docx" `
  --work-dir ".\qa-output" `
  --apply --yes --plan-apply --only-auto --render
```

`pipeline --render` writes `qa-output/rendered-before/` (before apply) and `qa-output/rendered/` (after apply). `--visual-repair` only sets `keepNext` and scales images to the column width.

Typical work dir:

```text
qa-output/
  before-audit.json
  after-audit.json          # --apply only
  document-model.json
  diff.json
  repair-plan.json
  apply.json
  validation.json
  visual-before.json
  visual.json
  visual-report.html
  report.html
  compliance.json
  pipeline-report.json
  rendered-before/          # --render
  rendered/                 # --render and --apply
```

Other commands:

```powershell
python .\standardize-docx-format\scripts\standardize.py model --input thesis.docx --output model.json
python .\standardize-docx-format\scripts\standardize.py repair-visual --input thesis.docx --output thesis.visual.docx --profile profile.json
```

## Thesis TOC safety

Chinese theses almost always have `toc 1` / `toc 2` lines that look like headings (`1 Introduction 1`). Number-stripping heading rules must exclude TOC styles:

```json
{
  "match": {
    "textRegex": "^\\d+\\s+\\S+",
    "styleNameNotRegex": "(?i)^toc",
    "textNotRegex": ".+\\d$",
    "inTextBox": false
  },
  "style": "Heading1",
  "textRegexReplace": {"pattern": "^\\d+\\s+", "replacement": ""},
  "maxMatches": 50
}
```

Matchers also include `textNotRegex`, `currentStyleNotIn`, `styleNameRegex`, and `styleNameNotRegex`. Display names are resolved from `word/styles.xml`. Example profiles already protect TOC styles. `thesis-cn.example.json` is a starting point, not a school-specific standard.

## Caption chapter numbers

Chinese heading numbers (`一、`) often must not appear in figure/table captions (`图 1-1`). Do not assume STYLEREF display text is what the school wants.

```json
{
  "captionNumbering": {
    "chapterStyle": "Heading 1",
    "headingDisplay": "chinese-counting",
    "chapterDisplay": "decimal",
    "separator": "-",
    "strategy": "styleref-arabic"
  }
}
```

| `strategy` | Meaning |
|---|---|
| `styleref-as-displayed` | Caption chapter number matches heading display |
| `styleref-arabic` | Headings may be Chinese; caption chapter is forced Arabic |
| `seq-chapter` | Independent `SEQ chapter` |
| `needs-confirmation` | Stay in `requirements.unresolved`; apply refuses by default |

Used when a `captionRule` omits `chapterFieldInstruction`.

## Minimal profile

See the Chinese README for a full JSON example, or start from [`references/thesis-cn.example.json`](standardize-docx-format/references/thesis-cn.example.json). Field docs:

- [`profile.schema.json`](standardize-docx-format/references/profile.schema.json)
- [`profile-schema.md`](standardize-docx-format/references/profile-schema.md)
- [`spec-driven-workflow.md`](standardize-docx-format/references/spec-driven-workflow.md)
- [`ooxml-advanced.md`](standardize-docx-format/references/ooxml-advanced.md)
- [`v2-architecture.md`](standardize-docx-format/references/v2-architecture.md)

## Placeholders, headers, paragraph rules

Placeholders may span runs and live in body, headers, footers, or text boxes. Build stable placeholders in a copy of the official cover; never flatten floating text boxes into body paragraphs.

`headersFooters[].type`: `default` | `first` | `even`. `action`: `set` | `clear` | `inherit`. `inherit` keeps the template’s headers/footers and does **not** drop refs. Creating `even` enables Word odd/even headers.

Bound `paragraphRules` with section, style, location, and `maxMatches`. Always exclude TOC when stripping numbering.

## Acceptance loop

```text
source / spec / template
  → Requirement IR + Document Spec
  → pre-edit audit + Semantic AST
  → Diff + Repair Plan
  → user confirmation
  → new DOCX
  → structural validate + Visual QA (optional PNG/PDF)
  → Word field refresh
  → final render review
```

Check cover alignment, abstract/TOC/body sections, Roman then Arabic page numbers, TOC styles vs headings, refreshed TOC page numbers, odd/even headers, overflow, orphan headings, caption splits, and leftover `{{PLACEHOLDER}}` text.

`validate` ok / `visual.ok` / a valid ZIP is structural only. `render.skipped` means nothing was rasterized.

## Limits

- Scripts create and preserve fields; they do not compute Word field results.
- TOC page numbers, PAGEREF, and NUMPAGES need Word (or a compatible renderer) to refresh.
- Exact covers depend on the template’s existing drawing coordinates.
- Word version, installed fonts, and printer drivers can shift pagination slightly.
- LibreOffice / `docx-cli` renders are approximate, especially for CJK numbering, STYLEREF, and TOC page numbers.
- Mixed Chinese headings vs Arabic caption chapters need an explicit `captionNumbering.strategy`.
- CITATION / BIBLIOGRAPHY fields are not a full reference manager.
- Macros, ActiveX, OLE, and vendor extensions need separate checks.
- Tests use synthetic OOXML, not anonymized school gold documents; CI does not run Word.
- Machine acceptance does not replace a visual pass.

## Data safety

Do not commit real theses, ID photos, signatures, student data, or unpublished templates. `.gitignore` drops DOCX and generated reports. Scripts do not upload documents. Inputs stay read-only; outputs are new paths. Scrub cover fields before sharing a profile.

## Develop and test

Compile list matches CI (includes `report.py`):

```powershell
python -B -X utf8 -m py_compile `
  .\standardize-docx-format\scripts\audit_docx.py `
  .\standardize-docx-format\scripts\apply_profile.py `
  .\standardize-docx-format\scripts\validate_docx.py `
  .\standardize-docx-format\scripts\document_model.py `
  .\standardize-docx-format\scripts\diff_profile.py `
  .\standardize-docx-format\scripts\repair_plan.py `
  .\standardize-docx-format\scripts\visual_qa.py `
  .\standardize-docx-format\scripts\layout_estimate.py `
  .\standardize-docx-format\scripts\visual_repair.py `
  .\standardize-docx-format\scripts\compile_requirements.py `
  .\standardize-docx-format\scripts\analyze_template.py `
  .\standardize-docx-format\scripts\standardize.py `
  .\standardize-docx-format\scripts\profile_schema.py `
  .\standardize-docx-format\scripts\plan_apply.py `
  .\standardize-docx-format\scripts\report.py
```

```powershell
python -B -X utf8 .\standardize-docx-format\scripts\test_paragraph_match_safety.py
python -B -X utf8 .\standardize-docx-format\scripts\test_v2_pipeline.py
python -B -X utf8 .\standardize-docx-format\scripts\profile_schema.py `
  --input .\standardize-docx-format\references\thesis-cn.example.json
python -B -X utf8 .\standardize-docx-format\scripts\profile_schema.py `
  --input .\standardize-docx-format\references\guangzhou-nanfang-thesis.example.json
```

## Design rules

1. Never overwrite the source; always write a copy.
2. Audit, then mutate, then validate.
3. Prefer style definitions over per-paragraph direct formatting.
4. Official covers/declarations win; headers/footers default to inherit.
5. Structure every natural-language rule; record uncertainty explicitly.
6. Keep fields as fields; never freeze displayed results as plain text.
7. Every semantic rule must be locatable, auditable, and preferably bounded; never remap TOC styles as headings.
8. A valid package is not a valid layout; render before sign-off.
9. Never apply a repair plan silently; `map_structure` and `pageBreakBefore` need explicit ids.

## Feedback

Issues welcome for anonymized school profiles, reproducible OOXML bugs, numbering/field/caption rules, Word vs LibreOffice mismatches, and PII-free minimal repros.

Possible follow-ups: deeper render-page inspection, screenshot OCR for the requirement compiler, a school profile library with gold theses, equation/footnote/table style models, a browser preview after the engine is stable, and renaming the `dist/` zip to match v2.1.

## License

No open-source license is attached. All rights reserved unless the copyright holder grants permission.
