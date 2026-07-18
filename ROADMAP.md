# Roadmap / 路线图

This roadmap communicates direction, not a promise of dates or delivery. Priorities may change based on verified user needs, maintenance capacity, security, and Office/Codex platform changes.

本路线图用于说明方向，不代表发布日期承诺；优先级会随真实需求、维护能力、安全因素及平台变化调整。

## Current focus / 当前重点

- Keep the skill contract conservative, non-destructive, and easy to audit.
- Expand structural tests for DOCX styles, sections, numbering, fields, and tables of contents.
- Improve bilingual documentation, contribution guidance, and reproducible issue reports.
- Keep the Word task pane compatible with its documented HTTPS adapter contract.
- Improve privacy and deployment guidance for institutions and individual users.

## Next / 下一阶段

- Publish synthetic, redistributable fixtures covering common Chinese and international thesis structures.
- Add regression tests for multilevel headings, Roman-to-Arabic page numbering, localized Word style IDs, and TOC refresh behavior.
- Improve task-pane accessibility, keyboard navigation, localization, progress reporting, and error recovery.
- Define a versioned JSON schema for format manifests and machine-readable change summaries.
- Provide reference deployment guidance or a minimal adapter example without operating a hosted service.
- Add automated release notes, checksums, and installation verification.

## Later / 后续探索

- Permission-reviewed institutional format profiles with provenance and version dates.
- Reusable policies for journals, reports, and other long-form Word documents.
- Optional local-first adapter patterns for privacy-sensitive environments.
- Broader interoperability tests across supported desktop and web Word environments.

## Explicit non-goals / 明确不做

- Hosting or silently routing user documents through a project-operated formatting backend.
- Guaranteeing acceptance by a university, journal, or review committee.
- Replacing the institution's current rules or final human review in Word.
- Generating deceptive academic work or bypassing academic-integrity requirements.

## Suggesting priorities / 提议优先级

Open a feature request describing the user problem, applicable formatting authority, privacy constraints, and a safe acceptance test. Votes and reactions help signal interest, but feasibility and safety remain part of prioritization.
