# Word Math Doctor

面向既有 Word 文档数学表达式的可复用 Agent Skill。它先做只读审计和作者裁定，再在明确授权的范围内修复为可编辑的 Word 原生 OMML；同时保护数学含义、正文、表格和版式。

当前版本：**v0.3.0-beta**。

## 适用范围

- 混合了 OMML、普通上下标、Unicode 伪公式、可见 LaTeX 或公式相关批注的 docx/docm；
- 论文返修、教材、讲义与技术报告中的公式规范化；
- 需要可复核的疑点清单、作者裁定和修改记录。

不用于推导或新写公式、公式图片 OCR、只改字体，或未经授权的全文改写。它不会猜测变量含义，也不会将某一份文稿的符号约定套用到另一份文稿。

## 安装

### 通用 Skills CLI

在目标项目目录中，任选你的 Agent 运行：

```bash
npx skills add lianxhcn/word-math-doctor --skill word-math-doctor --agent codex
npx skills add lianxhcn/word-math-doctor --skill word-math-doctor --agent claude-code
```

可先查看仓库中的可安装 Skill：

```bash
npx skills add lianxhcn/word-math-doctor --list
```

### Codex 插件市场

本仓库提供 Codex 市场清单：

```bash
codex plugin marketplace add https://github.com/lianxhcn/word-math-doctor.git --sparse .agents/plugins
```

随后在 Codex 中打开 `/plugins`，从市场安装 `word-math-doctor`，并在新会话中使用它。

### Claude Code 插件市场

```bash
claude plugin marketplace add lianxhcn/word-math-doctor
claude plugin install word-math-doctor@lianxhcn-agent-skills
```

不同 CLI 版本的 Agent 名称、目录和交互提示可能有所不同；以本机 CLI 显示的来源与安装目标为准。

## 调用示例

```text
$word-math-doctor 只读扫描 manuscript.docx 中的数学表达式，输出疑点清单，不修改原稿。

$word-math-doctor 仅处理 checklist 中已确认的公式项目，另存为 manuscript-revised.docx，并给出修改记录。

$word-math-doctor 对指定章节做 OMML 规范化；先确认范围、备份和环境能力，对含义不确定的项目保留原状。
```

## 工作层级

1. **结构诊断**：扫描并分类，不改原稿。
2. **作者裁定**：生成位置、证据、建议与裁定清单。
3. **限定修复**：只处理已授权且已确认的项目。
4. **高保证规范化**：在可写入 OMML、渲染、重新打开保存和独立复核的环境中进行。

扫描器为只读工具：

```bash
python skills/word-math-doctor/scripts/scan_docx_math.py manuscript.docx --output audits/manuscript-audit.json
```

默认不要公开审计 JSON、原稿、批注、路径或片段；这些材料可能泄露未公开文稿信息。

## 验证

```bash
python -m unittest discover -s tests -v
python skills/word-math-doctor/scripts/scan_docx_math.py --help
npx skills add . --list
```

## 许可

本项目采用 [CC BY-NC 4.0](LICENSE)。请保留署名、许可链接和修改说明；商业使用须先取得维护者许可。
