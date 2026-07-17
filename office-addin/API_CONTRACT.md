# Formatting backend API contract

The Word task pane is a client for a separately hosted HTTPS formatting service. The service adapts HTTP requests to the `format-thesis-docx` skill workflow and returns a new DOCX. The add-in never overwrites the active Word document.

## Security requirements

- Serve every endpoint over HTTPS.
- Configure CORS for the add-in origin, such as `https://localhost:3010` in development.
- Accept bearer authentication when the deployment is not strictly local.
- Never log manuscript contents, API tokens, or cover metadata by default.
- Delete temporary uploads and rendered QA artifacts according to the deployment's retention policy.
- Return a new file; never mutate a network path or active Word document in place.

## `GET /health`

Success response:

```json
{
  "status": "ok",
  "service": "format-thesis-docx-backend"
}
```

## `POST /v1/thesis/analyze`

Content type: `multipart/form-data`.

| Field | Type | Description |
| --- | --- | --- |
| `manuscript` | DOCX file | Temporary copy exported from the active Word document. |
| `authorities` | Repeated file | DOCX, DOTX, PDF, PNG, or JPEG formatting evidence. |
| `cover_metadata` | JSON string | Title, author, student ID, school, major, supervisor, degree, and date. |
| `options` | JSON string | TOC preference, output mode, safety settings, and client-side document signals. |

Example response with confirmation questions:

```json
{
  "analysis_id": "analysis_01J...",
  "status": "needs_confirmation",
  "summary": "Detected an official template and one conflicting PDF rule.",
  "questions": [
    {
      "id": "body_margin_left",
      "prompt": "Which left margin should be used for body pages?",
      "evidence": "The template uses 2.5 cm; the official PDF specifies 3.0 cm.",
      "proposed_default": "Use 3.0 cm from the higher-ranked written guide.",
      "impact": "The manuscript will repaginate.",
      "required": true,
      "options": [
        { "label": "3.0 cm", "value": "3.0", "recommended": true },
        { "label": "2.5 cm", "value": "2.5" }
      ]
    }
  ]
}
```

When there are no questions, return an empty `questions` array and a status such as `ready`.

## `POST /v1/thesis/format`

Accept the same multipart fields as the analyze endpoint, plus:

| Field | Type | Description |
| --- | --- | --- |
| `analysis_id` | String | Identifier returned by the analyze endpoint. |
| `confirmations` | JSON string | Map from confirmation question ID to the user's answer. |

Preferred success response:

- Status: `200 OK`
- Content type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- Header: `Content-Disposition: attachment; filename="thesis_formatted.docx"`
- Body: final DOCX bytes

The task pane also accepts a JSON response containing either:

```json
{ "download_url": "https://service.example/jobs/123/result", "filename": "thesis_formatted.docx" }
```

or:

```json
{ "file_base64": "UEsDB...", "filename": "thesis_formatted.docx" }
```

If formatting discovers a new consequential ambiguity, return `questions` again. The task pane pauses and resubmits after the user answers.

## Skill adapter behavior

The backend should:

1. Run `scripts/docx_profile.py` against the manuscript and DOCX/DOTX authorities.
2. Build the task-local thesis format manifest.
3. Apply the authority ranking and confirmation policy in `SKILL.md`.
4. Use the Codex `documents` capability for DOCX editing, dynamic TOC fields, and render QA.
5. Return only the final formatted DOCX and a concise machine-readable summary.
