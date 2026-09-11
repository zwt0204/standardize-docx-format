---
name: standardize-docx-format
description: Standardize, assemble, and audit Microsoft Word .docx files from an official reference template, a natural-language/PDF/web formatting specification, or a reusable JSON profile. Use for thesis/dissertation formatting, template filling, Chinese/Latin script fonts, semantic style mapping, Chinese or decimal multilevel numbering, per-section layout and page numbering, first/odd/even headers and footers, TOC/caption/cross-reference/bibliography fields, cover placeholders, structural and word-count validation, format cleanup, document QA, or batch standardization while preserving complex DOCX content.
---

# Standardize DOCX formatting

Normalize an existing `.docx` without rewriting its content model. Preserve the source, prefer style-level changes, apply advanced OOXML only through the bundled script, and verify the rendered result.

## Operating contract

- Treat the source document as immutable. Write to a new output path unless the user explicitly requests replacement.
- Require either a formatting specification, a reference template, or a JSON profile. If none exists, audit first and propose a profile instead of inventing a standard.
- Preserve text, images, equations, fields, comments, tracked changes, bookmarks, relationships, and unmodeled embedded objects unless the requested standard requires changing them.
- Prefer changing style definitions over applying direct formatting paragraph by paragraph.
- Exclude cover pages, declarations, approval forms, or other special sections when their layout intentionally differs.
- Never hard-code displayed field results as plain text. Preserve or create fields and request field refresh on open.



## Thesis / TOC safety

When standardizing Chinese undergraduate/graduate theses:

1. **Protect table-of-contents paragraphs.** Styles named like `toc 1` / `toc 2` must not be remapped by generic “starts with a number” heading rules.
2. **Prefer style-aware paragraph rules.** Use `styleNameNotRegex`, `styleNameRegex`, `currentStyleNotIn`, and `textNotRegex` before stripping numbering.
3. **Do not invent a school standard from a generic example.** `references/thesis-cn.example.json` is a starting point, not a school-specific final profile. If the user provides a school template or PDF rules, translate them into a profile first.
4. **Audit before/after headings and TOC.** Confirm TOC entries remain TOC styles and body chapters remain hierarchical headings.
5. **Field refresh.** After creating or preserving TOC fields, ask the user to update fields in Word on open.

Safer Arabic-number heading rule pattern:

```json
{
  "match": {
    "textRegex": "^\\d+\\s+\\S+",
    "styleNameNotRegex": "(?i)^toc",
    "textNotRegex": ".+\\d$"
  },
  "style": "Heading1",
  "textRegexReplace": {"pattern": "^\\d+\\s+", "replacement": ""},
  "maxMatches": 50
}
```


## Choose the execution path

1. **Reference template available:** use the template as the authoritative source. Preserve cover drawings/text boxes, sections, relationships, headers, footers, fields, and unknown extensions. Fill exact placeholders and map thesis content to semantic styles; do not rebuild encoded layout.
2. **Natural-language, PDF, web, or screenshot specification only:** read [references/spec-driven-workflow.md](references/spec-driven-workflow.md). Translate every measurable requirement into a v2 profile and put ambiguous or conflicting clauses in `requirements.unresolved`/`requirements.conflicts`. Do not apply until resolved unless the user explicitly accepts `--allow-unresolved`.
3. **Explicit JSON profile available:** read [references/profile-schema.md](references/profile-schema.md), audit the input, then apply and validate it.

Use `docx-cli` when the `docx` command exists. It is the preferred engine for locators, content-safe edits, validation, diffs, headers/footers, tables, and rendering. Run `docx <command> --help` before composing a mutation. Do not install missing dependencies without user authorization.

Use the bundled Python scripts for deterministic OOXML operations commonly missing from high-level tools: script-specific fonts, per-section layout, page-number restarts, semantic paragraph rules, Chinese/circled numbering, cross-run template replacements, independent first/default/even headers and footers, Word fields, bookmarks, captions, and measurable validation.

v2 compiler commands live in `scripts/standardize.py`. Prefer this sequence over jumping straight to `apply_profile.py`:

1. `audit` / `model` to understand the source document.
2. `compile` (spec text/PDF) or `analyze-template` (official DOCX) to draft a profile.
3. `diff` then `plan` so the user can see expected vs actual and confirm repairs.
4. `apply` only after confirmation. Never apply a repair plan silently. Prefer `apply --plan repair-plan.json --only-auto --yes` so only confirmed auto steps run. `map_structure` and `pageBreakBefore` require explicit `--only-ids`.
5. Validate the profile with `standardize.py schema` or by letting apply/compile/analyze-template run the bundled JSON Schema.
6. `validate` for structural checks and `visual` for layout risks / pagination estimates / optional render.
7. `repair-visual` only after confirmation, and only for conservative keepNext / image-scale fixes.

Read [references/v2-architecture.md](references/v2-architecture.md) when building a semantic AST, repair plan, or visual QA loop.

## Workflow

### 1. Prepare safely

- Resolve absolute input and output paths.
- Reject an output path equal to the input path.
- Use `<stem>.standardized.docx` as the default output name.
- Record any sections or styles that must remain untouched.

### 2. Audit before editing

Run:

```powershell
python scripts/audit_docx.py input.docx --output before-audit.json
```

When `docx-cli` is available, also run:

```powershell
docx validate input.docx --json
docx styles input.docx --used --json
docx outline input.docx --json
docx headers list input.docx
docx footers list input.docx
```

Render representative pages when Microsoft Word or LibreOffice is available. Inspect the cover, table of contents, first body page, a page containing a table or figure, and the final page.

Use the v2 audit to inspect bookmarks, fields, drawings/text boxes, content controls, numbering formats, section-to-header/footer relationships, and CJK/Latin text counts. A package that merely opens is not sufficient evidence of compliance.

### 3. Build the profile

- Read [references/profile-schema.md](references/profile-schema.md).
- Start from [references/thesis-cn.example.json](references/thesis-cn.example.json) for Chinese academic documents.
- For the audited Guangzhou Nanfang University adult-undergraduate template, start from [references/guangzhou-nanfang-thesis.example.json](references/guangzhou-nanfang-thesis.example.json), resolve its explicit open questions, and preserve the official DOCX as the base.
- Express measurements explicitly: inches for margins/indents, points for font/spacing, zero-based section indices for page numbering.
- Separate Latin, East Asian, and complex-script fonts.
- Represent heading numbering with composite patterns such as `%1`, `%1.%2`, `一、`, `（一）`, and `①` through `patterns` plus the corresponding Word number format.
- Keep semantic fields such as TOC, SEQ, REF, PAGEREF, CITATION, and BIBLIOGRAPHY as fields.
- Use `paragraphRules` only with bounded, auditable locators. Avoid broad global direct-format clearing.
- Use `replacements` for exact cover/header/footer/text-box placeholders. Never replace floating objects with ordinary body paragraphs.
- Record structural checks, abstract word counts, required fields/bookmarks, and residual placeholder patterns in `validation`.

### 3b. Explain before mutating

When the user has not already confirmed a profile application, run:

```powershell
python scripts/standardize.py diff --input input.docx --profile profile.json --text
python scripts/standardize.py plan --input input.docx --profile profile.json --text
```

Show the repair plan. Do not call `apply` until the user accepts it, or the request already amounts to an explicit apply.

### 4. Apply deterministic standardization

Run:

```powershell
python scripts/apply_profile.py `
  --input input.docx `
  --output output.standardized.docx `
  --profile profile.json
```

To apply a reviewed repair plan instead of the whole profile:

```powershell
python scripts/standardize.py apply `
  --input input.docx `
  --output output.standardized.docx `
  --profile profile.json `
  --plan qa-output/repair-plan.json `
  --only-auto --yes --force
```

The script applies package-safe OOXML changes and refuses in-place writes. It also refuses profiles with unresolved requirements/conflicts. Use `--force` only when replacing an existing output file is intentional; use `--allow-unresolved` only after the user explicitly accepts the risks.

When `docx-cli` is available, use its modeled verbs for requirements outside the profile, including table formatting, paragraph-scoped corrections, header/footer content, image sizing, and locator-based exceptions. Use `docx raw` only after reading [references/ooxml-advanced.md](references/ooxml-advanced.md), and only for constructs with no modeled command.

### 5. Verify after editing

Run:

```powershell
python scripts/audit_docx.py output.standardized.docx --output after-audit.json
python scripts/validate_docx.py --input output.standardized.docx --profile profile.json --output validation.json
```

When `docx-cli` is available, also run:

```powershell
docx validate output.standardized.docx --json
docx diff output.standardized.docx --against input.docx
docx render output.standardized.docx --out rendered-pages
```

Verify all of the following:

- The file opens and the ZIP package has no corrupt member.
- Required styles exist and are used by the intended paragraphs.
- Page size, orientation, margins, section count, page-number format, and page-number restart match the profile.
- Latin and East Asian fonts are distinct where required.
- Heading numbering remains automatic and subordinate levels restart correctly.
- Fields remain fields and `updateFields` is enabled when field results may be stale.
- Required bookmarks, TOC/SEQ/REF/PAGEREF/CITATION/BIBLIOGRAPHY fields, sections, chapter text, and word-count ranges pass the validation profile.
- No unresolved double-brace placeholder or template instructional text remains when prohibited.
- No unexpected reflow, clipped image, broken table, blank page, orphan heading, or cover-page shift appears in rendered pages.

## Standardization priorities

Apply changes in this order:

1. Requirement resolution and preservation/exclusion zones.
2. Global then per-section geometry and intentional section boundaries.
3. Default run properties, named style definitions, and semantic paragraph mapping.
4. Heading-linked multilevel numbering and manual-number cleanup.
5. Page numbering and first/default/even header/footer linkage.
6. Template replacements, body exceptions, captions, bookmarks, references, and fields.
7. Field refresh settings.
8. Validation, audit comparison, and Word render review.

Do not globally clear direct formatting unless the specification explicitly says every exception is invalid. Report residual direct formatting so a human or a targeted locator edit can resolve it safely.

## Resource routing

- Read [references/profile-schema.md](references/profile-schema.md) whenever creating or interpreting a standardization profile.
- Read [references/spec-driven-workflow.md](references/spec-driven-workflow.md) whenever no authoritative DOCX template exists.
- Read [references/ooxml-advanced.md](references/ooxml-advanced.md) before changing fields, numbering, script fonts, or section page numbering outside the bundled script.
- Use `scripts/audit_docx.py` for read-only package/style/layout inventory.
- Use `scripts/document_model.py` or `standardize.py model` for a semantic Document AST.
- Use `scripts/compile_requirements.py` to turn specification text into a draft profile; use `scripts/analyze_template.py` for an official template.
- Use `scripts/diff_profile.py` and `scripts/repair_plan.py` before any mutation.
- Use `scripts/profile_schema.py` or `standardize.py schema` to validate a profile against `references/profile.schema.json`.
- Use `scripts/apply_profile.py` for deterministic v1/v2 profile application, including `--plan` / `--only-auto` / `--only-ids`.
- Use `scripts/validate_docx.py` for structural, word-count, placeholder, field, bookmark, style, page-number, and header/footer checks.
- Use `scripts/visual_qa.py` for overflow, orphan headings, caption split, pagination estimates, and optional PDF/PNG render.
- Use `scripts/visual_repair.py` or `standardize.py repair-visual` for conservative keepNext / image-scale repairs.
- Use `scripts/standardize.py pipeline` to emit before/after audits, diff, repair plan, validation, visual reports, and `report.html`.

## Deliverables

Return:

- The standardized `.docx` output path.
- The applied profile path or a concise list of resolved rules.
- Before/after audit paths.
- Document AST / diff / repair-plan JSON when those steps ran.
- Machine-readable validation report and visual QA report.
- Any unresolved requirements that need Microsoft Word field refresh or visual judgment.
