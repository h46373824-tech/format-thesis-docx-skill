# Validation evidence

This page records demonstrated behavior without claiming universal university compliance.

## Latest desktop Word verification

- Date: 2026-07-18
- Platform: Microsoft Word desktop for Windows
- Document: synthetic Chinese master's-thesis demonstration

The demonstration contained fictional institutions, people, buildings and experiment values. It was created only to test document generation and formatting.

| Check | Result |
| --- | --- |
| Word opened without a repair prompt | Passed |
| Physical Word page count | 23 pages |
| OOXML sections | 3 |
| Page geometry | A4, 2.5 cm margins |
| Front matter | Uppercase Roman numbering |
| Main body | Decimal numbering restarted at 1 |
| Heading hierarchy | 37 paragraphs: 13 level-1, 18 level-2, 6 level-3 |
| Dynamic table of contents | Two Word pages, levels 1–3 |
| TOC page references | 37 `PAGEREF` fields |
| Data tables | 3 |
| Reference entries | 12 |
| External `altChunk` content | None |

## Verification method

1. Generate the DOCX from a structured manuscript and an explicit format manifest.
2. Inspect the ZIP/OOXML package for required parts, section properties, page size, margins, styles, numbering definitions, TOC fields, page fields, tables and references.
3. Open the file in desktop Word and accept the field-update operation.
4. Select the document, update the entire table of contents and save.
5. Review the cover, TOC, section transitions, body hierarchy, all tables, references, appendix and final page in Word.
6. Close Word and repeat the structural audit on the Word-saved file.

The final structural validator reported no failed checks. Word reported 23 pages, while the cached TOC correctly mapped front matter to Roman numerals and the body to Arabic page numbers.

## What this evidence does not prove

- It does not prove compatibility with every university template or every Word version.
- It does not validate academic claims, citations, originality or research ethics.
- It does not guarantee that a headless renderer and desktop Word will paginate identically.
- It does not test every possible equation editor, citation manager, macro or embedded object.
- It does not turn the synthetic demonstration into a submit-ready thesis.

## Recommended acceptance test for a new template

- Retain the original manuscript and create a formatted copy.
- Compare every required style against the official guide.
- Inspect all pages at 100% when a renderer is available.
- Open and save in the target Word version.
- Update the entire TOC and all fields.
- Verify cover metadata, section transitions, headers, footers and both page-number systems.
- Check tables, figures, equations, citations, references and intentional blank pages.
