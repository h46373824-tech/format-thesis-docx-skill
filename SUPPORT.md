# Support / 获取帮助

Community support is provided through GitHub Issues on a best-effort basis. 社区支持通过 GitHub Issues 提供，不承诺固定响应时间。

## Where to ask / 选择渠道

- **Reproducible defect / 可复现缺陷:** use the bug report template.
- **New capability or workflow / 新功能或工作流:** use the feature request template.
- **Security vulnerability / 安全漏洞:** follow [SECURITY.md](SECURITY.md); do not publish details in an issue.
- **Installation or usage question / 安装或使用问题:** open an issue with the `question` label if available, including your environment and the step that failed.

Before opening an issue, check the root README, `format-thesis-docx/SKILL.md`, and `office-addin/README.md`.

## Information to include / 请提供的信息

- operating system and version;
- Codex, Python, Node.js, Word, and browser versions when relevant;
- component and commit/version;
- exact command or user action;
- expected and actual behavior;
- sanitized logs and the smallest safe reproduction.

Remove names, student IDs, institution identifiers, document contents, tokens, private URLs, and other sensitive data. A synthetic document that reproduces the structure is much safer than a real thesis.

## Scope / 支持边界

The repository provides a Codex skill and a sideloadable Word task-pane client. It does not include a hosted formatting service, guaranteed institutional compliance, proofreading, academic evaluation, or emergency document recovery. Deployers are responsible for their own HTTPS adapter and data-handling policy.

本仓库不包含托管格式化服务，也不保证自动结果满足任何学校的最终审查要求。正式提交前请使用学校最新规范，并在 Microsoft Word 中人工复核。
