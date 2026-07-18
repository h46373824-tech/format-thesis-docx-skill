# Security Policy / 安全策略

## Supported versions / 支持范围

Security fixes target the current `1.0.x` release line and the default branch. 安全修复面向当前 `1.0.x` 版本线和默认分支。

## Reporting a vulnerability / 报告漏洞

Please use GitHub's private **Security > Report a vulnerability** flow for this repository. Include:

- the affected file, component, and commit or version;
- reproducible steps or a minimal proof of concept;
- expected impact and any known mitigations;
- whether the report may contain private document data.

请优先使用仓库的 GitHub 私密漏洞报告功能，并说明受影响组件、复现步骤、影响和缓解措施。不要在公开 Issue、日志或截图中放入论文正文、个人信息、凭证或真实服务地址。

If private vulnerability reporting is unavailable, open a minimal public issue requesting a private contact channel, but omit all exploit details and sensitive information.

## Response targets / 响应目标

Maintainers aim to acknowledge a complete report within 7 days and provide an initial assessment within 30 days. These are targets rather than a guaranteed service-level agreement. Please allow time for a coordinated fix before public disclosure.

维护者会尽力在 7 天内确认收到完整报告，并在 30 天内给出初步评估；该时间为目标而非服务等级承诺。公开披露前请预留协调修复时间。

## Security boundaries / 安全边界

- The Codex skill is designed to create a new output file instead of overwriting the manuscript.
- The Word task pane exports the document locally only when the user explicitly starts preview, analysis, or formatting. Only analysis and formatting send that exported copy to the configured backend.
- This repository does **not** provide a hosted formatting backend. Backend authentication, authorization, storage, retention, CORS, TLS, logging, and deletion policies belong to each deployer.
- Do not embed production secrets in task-pane JavaScript or commit them to this repository.
- University templates and thesis files may contain confidential, personal, or copyrighted material; use synthetic or explicitly authorized test data.

本项目不提供托管后端。部署者必须自行负责认证、授权、传输加密、跨域策略、日志脱敏、数据保留与删除。客户端代码中不得嵌入生产密钥。

Security reports are welcome for the skill instructions, DOCX profiler, Office Add-in client, manifest, dependency configuration, and public API contract. Vulnerabilities in a third-party deployment should also be reported to that deployment's operator.
