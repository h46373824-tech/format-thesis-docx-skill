# Thesis format manifest

Create `thesis_format_manifest.json` in the task's temporary workspace. It is an internal execution record, not a default deliverable.

## Shape

```json
{
  "version": 1,
  "manuscript": {
    "path": "thesis.docx",
    "output_path": "thesis_formatted.docx"
  },
  "authorities": [
    {
      "path": "university-template.docx",
      "kind": "official_template",
      "rank": 3
    }
  ],
  "rules": {
    "page.body.margin_left_cm": {
      "value": 3.0,
      "source": "format-guide.pdf",
      "location": "page 4",
      "confidence": "high",
      "requires_confirmation": false,
      "status": "pending",
      "qa": null
    },
    "toc.levels": {
      "value": [1, 2, 3],
      "source": "university-template.docx",
      "location": "existing TOC field",
      "confidence": "high",
      "requires_confirmation": false,
      "status": "pending",
      "qa": null
    }
  },
  "cover_fields": {
    "title": {
      "value": "Example thesis title",
      "target": "cover paragraph 3",
      "confidence": "medium",
      "confirmed": false
    }
  },
  "confirmations": [],
  "deviations": [],
  "qa": {
    "rendered_pages": 0,
    "all_pages_reviewed": false,
    "toc_checked": false,
    "page_numbering_checked": false
  }
}
```

## Rule keys

Use stable dotted keys. Recommended groups:

- `page.cover.*`
- `page.front_matter.*`
- `page.body.*`
- `section.*`
- `cover.*`
- `style.normal.*`
- `style.title.*`
- `style.heading1.*`
- `style.heading2.*`
- `style.heading3.*`
- `style.caption.*`
- `style.footnote.*`
- `style.references.*`
- `numbering.headings.*`
- `header.*`
- `footer.*`
- `page_number.front_matter.*`
- `page_number.body.*`
- `toc.*`
- `figure.*`
- `table.*`
- `references.*`
- `appendix.*`

## Confidence

- `high`: explicitly stated by an official authority or structurally consistent in an official template.
- `medium`: inferred from a consistent approved sample or several matching signals.
- `low`: visually guessed, structurally inconsistent, or based on a single ambiguous signal.

Set `requires_confirmation` independently of confidence. A medium-confidence reversible spacing value may proceed, while a medium-confidence cover-field mapping must be confirmed.

## Status

Use `pending`, `applied`, `skipped`, `blocked`, or `verified`.

For every applied rule, record a QA observation. Examples:

- `"matches template style Heading 1"`
- `"rendered without clipping on page 12"`
- `"Word TOC field includes levels 1-3"`
- `"user approved font substitution"`

