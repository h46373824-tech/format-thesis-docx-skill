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
| Sideload a Microsoft Word task pane add-in | Available in `office-addin/` |
| Format without an HTTPS skill-adapter backend | Not included |

This repository contains both the Codex skill and a sideloadable Microsoft Word task pane. Actual document editing and render verification are performed through Codex's `documents` capability.

The task pane exports the active DOCX, collects formatting evidence and confirmations, calls a configurable HTTPS backend, and downloads the result as a new file. The hosted formatting backend is deployment-specific and is not included. See [`office-addin/README.md`](office-addin/README.md) and [`office-addin/API_CONTRACT.md`](office-addin/API_CONTRACT.md).

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
office-addin/
  manifest.xml
  src/taskpane.html
  src/taskpane.css
  src/taskpane.js
  scripts/serve.mjs
  API_CONTRACT.md
```

## Validation

The skill has passed the Codex skill structure validator. The bundled DOCX profiler has also passed a smoke test covering Chinese heading detection, cover candidates, section margins, header extraction, and TOC-field detection.
