---
name: format-thesis-docx
description: Analyze and format academic thesis or dissertation Word files from uploaded university templates, approved samples, cover pages, screenshots, PDFs, or written formatting rules. Use when Codex must identify thesis formatting requirements, map cover metadata, normalize DOCX styles and heading hierarchy, configure sections and page numbering, insert or repair a Word table of contents, and pause for user confirmation only when a consequential requirement is ambiguous or conflicting.
---

# Format Thesis DOCX

Create a new, formatted `.docx` from a manuscript and the user's formatting evidence. Preserve the manuscript's academic content unless the user separately requests writing or editing.

Use the installed `documents` skill for DOCX editing, OOXML operations, rendering, and visual QA. Do not claim that this skill alone creates a Word ribbon add-in; it is the document-processing workflow that a future Office Add-in can call.

## Required inputs

Accept any combination of:

- The manuscript `.docx`.
- An official `.docx`/`.dotx` template or an approved thesis sample.
- A PDF, screenshot, or written university formatting guide.
- Cover-page metadata such as title, author, student ID, school, major, supervisor, degree, and date.

Require a manuscript before editing. Require at least one formatting authority, unless the user explicitly accepts a generic academic layout.

Treat formatting authorities in this order:

1. The user's explicit instruction for this run.
2. An official written university or journal guide.
3. An official Word template.
4. A formally approved sample thesis.
5. Inference from visual similarity or common academic conventions.

If higher-ranked sources conflict, apply the higher-ranked source. If equal-ranked sources conflict on a consequential item, confirm with the user.

## Safety rules

- Never overwrite the uploaded manuscript. Write a new file such as `<name>_formatted.docx`.
- Preserve wording, citations, equations, tables, figures, notes, and references unless a requested formatting change requires a structural adjustment.
- Preserve embedded objects and relationships. Do not rebuild complex equations, diagrams, or reference-manager fields from plain text.
- Record every user confirmation and every deliberate deviation from the supplied template.
- If a requested font is unavailable, do not silently substitute it when pagination could change.

## Workflow

### 1. Inventory and classify the evidence

Identify which file is the manuscript, which file controls appearance, and which files merely provide supporting rules. If file roles cannot be inferred safely, ask one concise question before editing.

For each `.docx` or `.dotx`, run:

```bash
python scripts/docx_profile.py input.docx --out profile.json
```

If the current environment exposes Python under another launcher or managed runtime, use that runtime. The profiler uses only the Python standard library.

For PDF or image instructions, extract only observable requirements. Do not infer hidden Word mechanics from appearance alone.

### 2. Build a format manifest

Create a task-local `thesis_format_manifest.json` using [format-manifest.md](references/format-manifest.md). For every rule, record:

- the intended value;
- the source and source location;
- confidence: `high`, `medium`, or `low`;
- whether user confirmation is required;
- the implementation method and later QA result.

Cover at least:

- page size, margins, orientation, sections, and page breaks;
- cover-page layout and metadata mapping;
- fonts for Latin text, East Asian text, and numbers;
- Normal, Title, Subtitle, Heading 1-3, captions, quotations, footnotes, and references;
- paragraph indentation, alignment, line spacing, and spacing before/after;
- heading hierarchy and multilevel numbering;
- abstract, keywords, acknowledgements, references, appendices, figures, and tables;
- headers, footers, page-number formats, and starting values;
- TOC title, location, included levels, indentation, leaders, and page-number display.

### 3. Run the confirmation gate

Read [confirmation-policy.md](references/confirmation-policy.md). Ask only about unresolved decisions that can materially affect meaning, institutional compliance, pagination, or irreversible layout.

Group related uncertainties into one short confirmation message. Include:

- the uncertain item;
- the evidence found;
- the proposed default;
- the effect of choosing differently.

Continue automatically for high-confidence, reversible formatting decisions.

### 4. Apply the template

Use the `documents` skill's template-following workflow. Treat the retained official template as the design authority.

Apply changes in this order:

1. Duplicate the manuscript to the output path.
2. Configure sections, page geometry, section breaks, headers, footers, and page-number restarts.
3. Reproduce or retain the official cover layout, then map confirmed cover metadata.
4. Normalize named Word styles instead of applying repeated direct formatting.
5. Detect true headings from existing styles, outline levels, numbering, wording patterns, and document structure.
6. Resolve uncertain heading levels before applying them in bulk.
7. Apply real Heading styles and real multilevel numbering. Do not use typed numbers as a substitute for numbering definitions.
8. Normalize captions, body paragraphs, quotations, notes, references, and appendices.
9. Preserve figures, tables, equations, citations, bookmarks, hyperlinks, and cross-references.

When a sample is visually authoritative but structurally poor, reproduce its visible rules with clean named styles instead of copying accidental direct-formatting defects.

### 5. Insert or repair the TOC

Make heading styles correct before creating the TOC.

Prefer a real Word TOC field for a Word deliverable. Include only confirmed heading levels, place it at the confirmed location, and mark fields to update when the document opens. If Word automation is available, update all fields before final save.

If a headless render cannot refresh TOC fields, use the `documents` skill's TOC and field-materialization tools for QA. Do not replace the final dynamic TOC with a static list unless the user approves that tradeoff.

Verify:

- no intended heading is missing;
- no body paragraph appears in the TOC;
- indentation and leader dots match the authority;
- page numbers match after the final pagination pass;
- the TOC title and its own page-number behavior are correct.

### 6. Render and inspect

Follow the `documents` render gate after every meaningful edit batch:

1. Render the DOCX to page PNGs.
2. Inspect every page at 100% zoom.
3. Fix layout defects and re-render.

Check cover alignment, blank pages, widows/orphans, heading placement, table and figure clipping, equation integrity, mixed fonts, section transitions, headers/footers, Roman-to-Arabic page-number changes, TOC pagination, and reference indentation.

Treat odd/even pagination as a Word-specific risk. When `evenAndOddHeaders` is enabled and a later `nextPage` or `oddPage` section restarts at an odd number such as 1, Word may insert a virtual blank even page that is not obvious in a headless render. Review `docx_profile.py` warnings and verify the physical page count in desktop Word. If the authority does not require distinct odd/even headers or footers, disable the setting and remove `even` references; otherwise preserve it and confirm that the inserted parity page is intended.

After Word saves a document, validate styles semantically rather than requiring literal English style IDs. Word may localize or shorten `styleId` values while retaining the style name, `outlineLvl`, numbering, and paragraph references. It may also omit the default `decimal` page-number format and remove `updateFields` after successfully refreshing a dynamic TOC. Accept those normalizations only when the section start value, TOC field, PAGEREF cache, page metadata, and rendered result remain consistent.

If LibreOffice is unavailable, perform structural OOXML audits and disclose that visual QA could not be completed.

### 7. Deliver

Return:

- the formatted `.docx`;
- a concise list of applied formatting rules;
- confirmations and assumptions that affected the result;
- any unresolved limitation, such as a field that Word must refresh on open.

Do not return profiles, rendered PNGs, or the manifest unless the user asks for them.

## Word Add-in handoff

If the user wants an actual button inside desktop or web Word, treat that as a separate Office Add-in project. Use this skill as the processing contract:

- input: manuscript, authority files, metadata, and prior confirmations;
- interaction: return batched confirmation questions when required;
- output: a new formatted DOCX plus a structured change summary;
- rule: never modify the active document destructively without an explicit save-copy choice.
