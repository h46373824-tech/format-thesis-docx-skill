# Validation evidence

This page records what the repository can re-run. It does not claim that a finished thesis matches a university's rules, and it does not claim that desktop Word pagination was verified for the current revision.

## Automated checks

GitHub Actions and the local commands in the README run:

- `python -m unittest discover -s tests -v`
- `node scripts/test-client-policy.mjs` from `office-addin/`
- skill metadata checks, Python compilation, and Office manifest validation

`tests/test_docx_profile.py` builds synthetic DOCX packages in a temporary directory. The packages are not opened in Microsoft Word. They check that the profiler:

- resolves theme fonts and `basedOn` inheritance, including a Chinese default style ID;
- keeps Heading 1 character styles out of the heading map;
- reads direct run formatting, content controls, table paragraphs, and OMML text;
- reports `firstLineChars`, 1.5 line spacing, mirror margins, and document grid;
- treats outline level 9 as body text and does not let a `Heading 10` name override outline level 8;
- stops cover candidates before the first section break;
- reads simple-field TOC and PAGE instructions and numbering definitions;
- ignores `numId` 0;
- keeps only one AlternateContent branch;
- continues when a header part is not XML;
- rejects DTD entity declarations and over-limit packages;
- reports a VBA part inside a `.docx` separately from a macro-enabled extension.

The task-pane tests check that a download URL must be HTTPS on the configured service host, that cross-host and credentialed URLs are rejected, that redirects are followed only on that same host, and that saved names end in `.docx`.

## Historical desktop Word note

On 2026-07-18 a synthetic Chinese thesis was opened in desktop Word for Windows during early development. That file is not stored in this repository, and the session was not repeated for later commits. Notes from that session mentioned 23 physical pages, 3 sections, 37 headings, and 37 `PAGEREF` fields. Those numbers are not a current test result. Do not describe the current revision as pagination-verified or as producing a submission-ready thesis unless someone repeats the check in the target Word version and records it separately.

## What this evidence does not prove

- Compatibility with every university template or Word version.
- Academic claims, citations, originality, or research ethics.
- That a headless renderer and desktop Word paginate the same way.
- Equation editors, citation managers, macros, or embedded objects beyond the synthetic cases above.
- That the Word task pane formats a document by itself. Formatting still requires the skill workflow and a separately deployed adapter.

## Recommended acceptance test for a new template

- Retain the original manuscript and create a formatted copy.
- Compare required styles with the official guide.
- Inspect pages in the target Word version when a visual check is required.
- Update the table of contents and fields in Word.
- Verify cover metadata, section breaks, headers, footers, and both page-number systems.
- Check tables, figures, equations, citations, references, and intentional blank pages.
