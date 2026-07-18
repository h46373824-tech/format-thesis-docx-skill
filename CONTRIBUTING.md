# Contributing / 参与贡献

Thank you for helping improve `format-thesis-docx`. Contributions in English or Chinese are welcome. 感谢你帮助改进本项目；中文和英文贡献都欢迎。

## Before you start / 开始之前

- Search existing issues before opening a new one. 提交前请先搜索现有 Issue。
- Use the bug or feature template and keep one topic per issue. 请使用对应模板，并让每个 Issue 聚焦一个主题。
- For a substantial workflow, API-contract, or compatibility change, open an issue before investing in an implementation. 对工作流、接口契约或兼容性有较大影响的改动，请先发 Issue 讨论。
- Never upload a real thesis, personal information, access token, proprietary university template, or copyrighted sample without permission. 请勿上传真实论文、个人信息、令牌、未获授权的学校模板或受版权保护的样例。

## Project boundaries / 项目边界

This repository contains:

- a Codex skill that analyzes formatting evidence and coordinates safe DOCX formatting;
- a standard-library Python DOCX profiler;
- a sideloadable Word task-pane client and its public API contract.

It does **not** include or operate a hosted formatting backend. The task pane must be connected to a separately deployed HTTPS adapter. A contribution must not imply that user documents are automatically uploaded to a service maintained by this project.

本仓库包含 Codex Skill、DOCX 分析脚本、Word 任务窗格客户端及公开接口契约；**不包含也不运营托管格式化后端**。任务窗格需要连接到部署者自行维护的 HTTPS 适配服务，任何贡献都不应暗示文档会自动上传到本项目运营的服务。

## Development setup / 开发环境

Requirements:

- Python 3.10 or later for `docx_profile.py`;
- Node.js 20 or later for the Office Add-in;
- Microsoft Word only when manually testing sideloading or Word-specific behavior.

Run the local checks from the repository root:

```powershell
python -m compileall -q format-thesis-docx/scripts
python format-thesis-docx/scripts/docx_profile.py --help

cd office-addin
npm ci
npm run check
npm run validate
```

The profiler uses only the Python standard library. Office Add-in dependencies are pinned by `office-addin/package-lock.json`.

## Making a change / 提交修改

1. Fork the repository and create a focused branch.
2. Keep behavior aligned with `format-thesis-docx/SKILL.md`, especially non-destructive output, evidence ranking, confirmation gates, and render/Word QA.
3. Update documentation or `office-addin/API_CONTRACT.md` when observable behavior changes.
4. Add only synthetic, anonymized fixtures. Do not commit generated DOCX/PDF files containing user material.
5. Run all relevant checks and describe the result in the pull request.

请保持改动范围清晰；行为修改必须继续遵守“另存新文件、证据优先级、关键歧义确认和版式验证”等安全规则。测试材料必须是合成或彻底匿名化的数据。

## Pull requests / 拉取请求

A useful pull request includes:

- the problem and intended outcome;
- the affected component (`skill`, `profiler`, `office-addin`, or documentation);
- tests or manual verification performed;
- screenshots only when they are sanitized and materially help review;
- compatibility or security considerations.

Small, reviewable pull requests are preferred. Maintainers may ask to split unrelated changes. By contributing, you agree that your contribution is licensed under the repository's MIT License.

## Reporting security issues / 安全问题

Do not disclose a vulnerability in a public issue. Follow [SECURITY.md](SECURITY.md). 请勿在公开 Issue 中披露漏洞细节，请按安全策略私下报告。
