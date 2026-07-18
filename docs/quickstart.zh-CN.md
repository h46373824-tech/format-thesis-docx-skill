# 中文安装与使用指南

## 适用场景

当你已经有论文 Word 原稿，并且拥有学校模板、PDF 规范、封面样例、截图或文字格式要求时，可以使用本 Skill 自动整理版式。它不会替你证明论文内容正确，也不会把虚构数据变成真实研究结果。

## 安装 Codex Skill

### Windows PowerShell

```powershell
git clone https://github.com/h46373824-tech/format-thesis-docx-skill.git
New-Item -ItemType Directory -Force "$HOME\.codex\skills" | Out-Null
Copy-Item -Recurse -Force ".\format-thesis-docx-skill\format-thesis-docx" "$HOME\.codex\skills\"
```

### macOS 或 Linux

```bash
git clone https://github.com/h46373824-tech/format-thesis-docx-skill.git
mkdir -p ~/.codex/skills
cp -R format-thesis-docx-skill/format-thesis-docx ~/.codex/skills/
```

如果 Codex 没有立即显示新 Skill，请重启 Codex，再检查技能列表中是否出现 `format-thesis-docx`。

## 推荐输入

至少提供：

1. 论文原稿 `.docx`；
2. 一份格式依据，例如官方模板、PDF 规范、正式样例或明确的文字要求。

建议同时提供：

- 论文题目、作者、学号、学院、专业、导师、学位类型和日期；
- 需要保留的参考文献管理器字段、公式、图表或嵌入对象说明；
- 是否允许 Word 在目录前后插入空白页；
- 学校要求的特殊字体及其是否已经安装。

## 推荐提示词

```text
使用 $format-thesis-docx 按我上传的学校模板整理这篇论文。
不要改写正文内容，不要覆盖原文件；请创建新的 DOCX，规范封面、正文、
三级标题、分节、页眉页脚和页码，并添加可更新的 Word 动态目录。
只有遇到会影响学校合规性、内容含义或分页的关键歧义时再向我确认。
```

## 工作过程

1. 区分论文原稿与格式依据。
2. 提取页边距、字体、标题、封面、分节、页码和目录规则。
3. 建立本次任务的格式清单并标明规则来源。
4. 将真正影响结果的歧义集中成少量确认问题。
5. 生成新文件，保留正文、引用、公式、表格、图片与文档关系。
6. 更新动态目录和页码字段。
7. 进行 OOXML 结构检查，并在环境允许时渲染或用桌面版 Word 验证。

## Word 任务窗格

仓库内的 `office-addin/` 是可以侧载的 Microsoft Word Office Add-in 客户端。

```powershell
cd format-thesis-docx-skill\office-addin
npm ci
npm run validate
npm start
```

要求 Node.js 20 或更高版本。首次运行可能会安装本地 HTTPS 开发证书。

任务窗格负责导出当前文档、收集模板和封面信息、显示确认问题，并下载格式化后的副本。它本身不包含 AI 排版后端；必须配置符合 [`API_CONTRACT.md`](../office-addin/API_CONTRACT.md) 的 HTTPS 服务。服务地址被有意隐藏在任务窗格界面之外。

## 常见问题

### 为什么必须生成新文件？

论文往往包含域、书签、公式、批注和嵌入对象。保留原稿可以在分页或兼容性出现问题时安全对照和回退。

### 为什么目录打开后还需要更新？

Word 的页码由实际排版环境计算。最终保存前应在 Word 中选择整个目录并执行“更新整个目录”，再检查罗马页码和正文页码是否正确衔接。

### 没有学校模板能否使用？

可以，但需要你明确接受通用学术排版。通用格式不能替代学校的正式要求。

### 会不会上传论文到网络？

Codex Skill 的实际数据路径取决于你使用的 Codex 环境。Word 任务窗格只有在你配置 HTTPS 适配器后才会向该服务发送文档。部署者应在生产环境中说明数据保存、访问控制和删除策略，不应在前端代码中放置令牌。

### 是否支持 `.doc`？

推荐先在 Microsoft Word 中将旧版 `.doc` 另存为 `.docx`，再进行格式分析。

## 提交问题

报告问题时请提供脱敏后的最小示例、学校格式来源、Word 版本、操作系统、预期结果和实际结果。不要公开上传含真实个人信息、未发表研究数据或受限模板的论文。
