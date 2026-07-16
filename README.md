# Format Thesis DOCX Skill

A Codex skill for analyzing university thesis templates and formatting Word manuscripts.

## Capability status

| Capability | Status |
| --- | --- |
| Analyze DOCX/DOTX formatting, styles, sections, headings, headers, footers, cover candidates, and TOC fields | Available |
| Apply thesis formatting through Codex's `documents` capability | Available |
| Normalize headings, page geometry, page numbering, cover content, and table of contents | Available |
| Pause for confirmation when formatting evidence is ambiguous or conflicting | Available |
| Install as a Codex skill | Available |
| Install directly as a Microsoft Word ribbon add-in | Not included |

This repository contains a Codex skill, not a standalone DOCX converter and not an Office Add-in. Actual document editing and render verification are performed through Codex's `documents` capability.

The skill includes a Word Add-in handoff contract so it can serve as the document-processing workflow behind a separately developed Microsoft Office Add-in. A Word ribbon button, task pane, Office manifest, and hosted add-in service are not included in this repository.

## Install

Copy the `format-thesis-docx` directory into:

```text
~/.codex/skills/format-thesis-docx
```

Then invoke:

```text
$format-thesis-docx
```

Example:

```text
Use $format-thesis-docx to format my thesis using the uploaded university
template, add a dynamic table of contents, and ask me only about consequential
ambiguities.
```

## Expected inputs

- A thesis or dissertation manuscript in `.docx` format.
- An official `.docx`/`.dotx` template, approved sample, PDF guide, screenshot, or written formatting rules.
- Cover metadata when it is not already present in the manuscript.

## Safety behavior

- Creates a new output file instead of overwriting the manuscript.
- Preserves academic content, citations, equations, tables, figures, and document relationships.
- Uses real Word heading styles and a real TOC field.
- Records material assumptions and asks for confirmation before consequential ambiguous changes.
- Requires render-and-inspect QA through the Codex `documents` skill.

## Repository layout

```text
format-thesis-docx/
  SKILL.md
  agents/openai.yaml
  references/confirmation-policy.md
  references/format-manifest.md
  scripts/docx_profile.py
```

## Validation

The skill has passed the Codex skill structure validator. The bundled DOCX profiler has also passed a smoke test covering Chinese heading detection, cover candidates, section margins, header extraction, and TOC-field detection.

