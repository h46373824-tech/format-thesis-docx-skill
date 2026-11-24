# Changelog / 更新日志

All notable changes to this project are documented here. This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and intends to use [Semantic Versioning](https://semver.org/) for releases.

所有重要变更都会记录在此；项目计划按语义化版本发布。

## Unreleased

### Fixed / 修复

- DOCX profiler now resolves theme fonts and `basedOn` inheritance, including Chinese style IDs, and reports direct formatting, content controls, table paragraphs, formulas, character indents, document grid, and mirror margins.
- Heading detection treats outline level 9 as body text, keeps numbering and simple-field TOC/PAGE instructions, and stops cover extraction at the first heading, page break, or section break.
- AlternateContent contributes one branch. Unreadable header/footer parts degrade to warnings. DTD entities and oversized packages are rejected.
- Task pane downloads stay on the configured HTTPS origin, follow only same-origin redirects, force a `.docx` base name, and require a ZIP signature.
- Manifest permission is `ReadDocument`. The skill text no longer describes the Word task pane as a future add-in.
- Validation docs now separate re-runnable synthetic tests from the unreproduced 2026-07-18 desktop Word note.

### Security / 安全

- Development dependencies are now `office-addin-debugging` 7.0.1, `office-addin-manifest` 3.0.1, and `office-addin-dev-certs` 3.0.1. Direct `adm-zip` is 0.6.1. `npm audit` still reports 2 high and 3 moderate issues through the Office toolkit peer dependency (`@azure/msal-node`, `glob` / `minimatch`). Those packages are not part of the task pane runtime, and `npm audit fix` was not applied.

## 1.0.0 - 2026-07-18

### Added / 新增

- Initial `format-thesis-docx` Codex skill workflow.
- Formatting-evidence precedence and confirmation policy.
- Thesis format-manifest reference.
- Standard-library DOCX profiler for styles, sections, headings, headers, footers, cover candidates, and TOC fields.
- Sideloadable Microsoft Word task-pane client.
- Public HTTPS formatting-adapter contract.
- Bilingual contribution, support, security, conduct, citation, and roadmap documentation.
- Structured GitHub issue forms and pull request template.
- GitHub Actions validation for skill metadata, Python syntax, and the Office Add-in.
- Bilingual quick-start and validation evidence, plus GitHub promotional artwork.

### Security / 安全

- Non-destructive output requirement: formatting produces a new DOCX rather than overwriting the manuscript.
- Guidance for protecting manuscript content, credentials, and deployment-specific service configuration.
