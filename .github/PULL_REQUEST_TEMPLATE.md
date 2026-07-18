## Summary / 摘要

<!-- What problem does this PR solve? 这个 PR 解决什么问题？ -->

## Changes / 改动

<!-- List the observable behavior and affected components. 列出可观察行为及受影响组件。 -->

## Validation / 验证

<!-- Include commands, results, and any sanitized manual Word checks. -->

- [ ] Skill metadata and instructions remain valid / Skill 元数据和说明有效
- [ ] `python -m compileall -q format-thesis-docx/scripts`
- [ ] `npm ci` in `office-addin/`
- [ ] `npm run check` in `office-addin/`
- [ ] `npm run validate` in `office-addin/`
- [ ] Documentation/API contract updated when behavior changed / 行为变化时已更新文档或接口契约

## Safety and privacy / 安全与隐私

- [ ] The change preserves non-destructive output unless explicitly documented and approved / 改动保持非破坏性输出
- [ ] Fixtures, logs, and screenshots are synthetic or fully anonymized / 测试材料、日志和截图均为合成或彻底匿名化
- [ ] No secret, private URL, thesis content, or unauthorized template is included / 不包含密钥、私有地址、论文正文或未授权模板
- [ ] The change does not imply that this repository includes or operates a hosted formatting backend / 不暗示本仓库包含或运营托管格式化后端

## Compatibility and risks / 兼容性与风险

<!-- Note Word/Office.js, Python, Node.js, API-contract, pagination, or migration risks. -->

## Related issue / 关联 Issue

<!-- Example: Closes #123 -->
