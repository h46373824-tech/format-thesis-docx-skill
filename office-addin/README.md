# Microsoft Word task pane

This directory contains a sideloadable Microsoft Office Add-in task pane for Word.

## What works in the task pane

- Export the active Word document as a temporary DOCX copy.
- Inspect paragraph count, heading-style candidates, selection text, and export size.
- Collect university templates, approved samples, PDF/image rules, and cover metadata.
- Send the manuscript and formatting evidence to an internally configured HTTPS backend.
- Render consequential confirmation questions returned by the skill adapter.
- Download the formatted result as a new DOCX without overwriting the active document.

The task pane is not the formatting engine. Deploy an HTTPS adapter that implements [API_CONTRACT.md](API_CONTRACT.md) and invokes the `format-thesis-docx` skill workflow.

## Requirements

- Microsoft Word with Office Add-ins enabled.
- Node.js 20 or later.
- An HTTPS formatting backend for analysis and formatting operations.

## Install dependencies

```powershell
cd office-addin
npm install
```

## Validate the manifest and JavaScript

```powershell
npm run validate
npm run check
```

## Start and sideload in desktop Word

```powershell
npm start
```

The first run installs and trusts a localhost development certificate. Word should open with the add-in sideloaded. If automatic sideloading is blocked by local Office policy, start the HTTPS server separately:

```powershell
npm run dev-server
```

Then configure a trusted network-share catalog or use your organization's Office Add-in catalog with `manifest.xml`.

## Sideload in Word on the web

1. Run `npm run dev-server`.
2. Open a document in Word on the web.
3. Open **Home > Add-ins > More Settings > Upload My Add-in**.
4. Upload `manifest.xml`.

## Backend configuration

The service address is intentionally hidden from the task-pane UI. Development builds use the `SERVICE_CONFIG.baseUrl` value near the top of `src/taskpane.js`, which defaults to `https://localhost:8787`.

Do not embed production bearer tokens or other secrets in task-pane JavaScript. For production, place authentication and the real backend address behind a same-origin server-side proxy, then point `SERVICE_CONFIG.baseUrl` at that proxy origin.

## Production deployment

For production, host the files in `src/` on an HTTPS origin and replace every `https://localhost:3010` value in `manifest.xml` with that origin. Validate the edited manifest before distributing it through an organizational catalog or Microsoft Marketplace.
