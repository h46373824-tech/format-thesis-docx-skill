# Confirmation policy

Use this policy after extracting the formatting evidence and before bulk editing.

## Confirm before proceeding

Ask for confirmation when any of these conditions applies:

- Two authorities of equal rank specify different page size, margins, fonts, heading numbering, page numbering, TOC depth, or cover layout.
- The cover contains unlabeled values that cannot be mapped confidently to author, student ID, major, supervisor, degree, school, or date.
- More than one plausible paragraph can be the thesis title, author, abstract heading, first chapter, references heading, or appendix boundary.
- Heading detection would reclassify many paragraphs or change numbering across the document.
- A font substitution could alter pagination or violate an explicit institutional requirement.
- The reference uses different rules for front matter and body matter, but the correct section boundary is unclear.
- The output would remove, flatten, or rebuild equations, citations, content controls, tracked changes, comments, macros, embedded files, or reference-manager fields.
- The user has not indicated whether to keep existing tracked changes or accept them.
- The requested behavior would overwrite the original file.

## Proceed automatically

Do not interrupt the user for:

- High-confidence style values directly stated in an official guide.
- Page geometry and named-style values consistently present in an official template.
- Reversible cleanup such as removing accidental direct formatting after the matching named style is known.
- Inserting a dynamic TOC after heading levels and TOC placement are unambiguous.
- Creating a new output copy instead of modifying the original.
- Updating fields, bookmarks, and internal links when the visible result is unchanged except for correctness.

## Default decisions

Use these defaults only when no authority specifies otherwise:

- Preserve all academic content.
- Create a new file with `_formatted` appended to the name.
- Use three TOC levels.
- Use real Heading 1-3 styles and real multilevel numbering.
- Keep the cover unnumbered.
- Keep preliminary pages separate from the Arabic-numbered body.
- Mark fields to update on open.
- Preserve citation and equation fields.

## How to ask

Batch related questions. Avoid asking the user to choose Word implementation details they should not need to know.

Good:

> 模板示例的正文页边距是 2.5 cm，但学校 PDF 规定左侧 3.0 cm。PDF 的优先级更高，我建议采用左 3.0 cm、其余 2.5 cm；这会让正文重新分页。是否按 PDF 执行？

Bad:

> w:pgMar 应该设置成多少 twips？

After confirmation, store the answer in the manifest and do not ask the same question again.

