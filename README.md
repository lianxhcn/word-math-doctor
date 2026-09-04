# Word Math Doctor

一个可复用的 Codex Skill，Skill ID 为 word-math-doctor，用于扫描、裁定并修复已有 Word 文档中的数学表达式。
目标是将应处理的公式变为可编辑的 Word 原生 OMML，同时保护数学含义、正文、表格、
图片和整体版式。

它适合论文返修、教材、讲义和技术报告；它不是 Word 宏、Office 加载项或自动代写工具。
在没有足够公式依据或作者授权时，它不会猜测变量含义，也不会把公式图片静默替换为
原生公式。

> **公开 beta：v0.2.0-beta。**本版本提供可复核的工作流和只读 OOXML 扫描器；它不内置
> Word 自动化转换器。只有具备 Word/OMML 写入、渲染和重新保存能力的 Agent 环境，才能完成
> 已授权的公式修复与全文验收。

## 版本状态与能力边界

- **可以可靠完成**：结构扫描、疑点分类、作者裁定清单、修改记录和转换前后计数比较。
- **需要相应环境才可完成**：写入真实 OMML、逐页渲染、Microsoft Word 重新打开保存和独立复核。
- **当前尚未完成**：跨 Agent 的统一 benchmark、多平台渲染兼容性结论和稳定版 v1.0.0 声明。

## 先看这三个要点

1. 推荐用 Skills CLI 安装。它会按照所选的 Agent 和安装范围放置文件，不需要手工猜测
   本机目录。
2. 第一次使用应先做只读扫描，再由作者确认有歧义的公式，最后才执行转换。
3. GitHub 公开仓库即可供读者直接安装；skills.sh 是可选的发现渠道，不是使用本 Skill
   的前提。

## 适合解决什么问题

- 文稿中混有 Word 原生公式、普通上标/下标、Unicode 伪公式和可见 LaTeX；
- 公式显示为 Cambria Math，但实际仍是普通文本；
- Word 批注要求统一修复上下标、符号、公式编号或正斜体；
- 已有 OMML 公式需要检查复合下标、函数正体、定界符、花体和实际渲染；
- 需要保留正文、表格、图片、页眉页脚和分页，并交付可审计的修改记录。

## 不做什么

- 不把公式转成图片，也不用字体伪装为公式；
- 不根据猜测补写或改变数学含义；
- 不在未经授权的情况下删除批注、修订痕迹或重建公式编号；
- 不将图中的公式自动 OCR 后替换为原生公式。

## 快速安装：推荐使用 Skills CLI

### 1. 查看仓库中可安装的 Skill

先确保本机已安装 Node.js，因此可使用 npx。下面的命令只列出仓库内容，不安装：

~~~bash
npx skills add lianxhcn/word-math-doctor --list
~~~

### 2. 安装到当前项目的 Codex

在需要处理 Word 文档的项目目录中运行：

~~~bash
npx skills add lianxhcn/word-math-doctor --skill word-math-doctor --agent codex
~~~

该命令会在确认前展示安装范围。首次安装建议保留交互确认，以便检查来源和目标位置。
只有在自动化环境中已完成审查时，才附加 --yes。

### 3. 可选：安装到本机全部 Codex 项目

若希望之后所有本机项目都能使用，再运行：

~~~bash
npx skills add lianxhcn/word-math-doctor --skill word-math-doctor --agent codex --global
~~~

### 4. 验证并调用

~~~bash
npx skills list --agent codex
~~~

重新打开 Codex 后，可先输入 /skills 确认它已被发现，再显式调用：

~~~text
$word-math-doctor 先扫描这份 Word 文档中的数学公式问题，不修改原稿。
~~~

### 5. 更新

在原安装范围内运行：

~~~bash
npx skills update word-math-doctor
~~~

如果 CLI 报告的 Agent 名称或安装位置与本机 Codex 版本不一致，以 CLI 的提示为准；
这是避免手动复制到错误目录的主要原因。

## 手动安装（没有 Skills CLI 时）

独立 Skill 目录是 skills/word-math-doctor。克隆仓库后，将该目录放入目标项目的
.agents/skills 下，再重新启动 Codex：

~~~bash
git clone --depth 1 https://github.com/lianxhcn/word-math-doctor.git
mkdir -p .agents/skills
cp -R word-math-doctor/skills/word-math-doctor .agents/skills/
~~~

Windows PowerShell 可使用：

~~~powershell
git clone --depth 1 https://github.com/lianxhcn/word-math-doctor.git
New-Item -ItemType Directory -Force .agents/skills | Out-Null
Copy-Item -Recurse word-math-doctor/skills/word-math-doctor .agents/skills/
~~~

目前的 OpenAI 文档说明，Codex 可从项目或用户级的 Agent Skills 目录发现 Skill。不同版本
的本机工具对手动目录的兼容性可能不同；若没有被发现，请改用上面的 Skills CLI 路径。

## 实操调用示例

| 目的 | 直接粘贴到 Codex 的请求 |
| --- | --- |
| 只读诊断 | $word-math-doctor 扫描 manuscript.docx 中所有疑似公式问题，输出 audit.json 和待作者确认清单，不修改原稿。 |
| 根据裁定修复 | $word-math-doctor 依据 decision-checklist.docx 中已确认的项目修复 manuscript.docx，另存为 manuscript-revised.docx，并报告每一项修改。 |
| 批注返修 | $word-math-doctor 只处理 manuscript.docx 中未解决批注所指的公式；对数学含义不确定的项目保留原状并列入待确认项。 |
| 全文规范化 | $word-math-doctor 对 manuscript.docx 做全文 OMML 规范化。先给出范围、能力检查和备份方案；完成后提供结构审计、渲染验收、重新保存和独立复核结果。 |

全文规范化属于高风险操作。开始前应明确授权处理范围，保留原件，并为有歧义的表达式提供
作者裁定。完成一个 Markdown checklist 不等于已经生成或验证了 Word 原生公式。

## 只读结构扫描器

仓库内附 Python 标准库实现的 OOXML 扫描器。它不改写文档，可用于转换前后基线对比：

~~~bash
python skills/word-math-doctor/scripts/scan_docx_math.py manuscript.docx --output audits/manuscript-audit.json
python skills/word-math-doctor/scripts/scan_docx_math.py manuscript-revised.docx --baseline audits/manuscript-audit.json --output audits/revised-audit.json
~~~

扫描器识别 OMML、普通上下标、Unicode 伪公式、可见 LaTeX、批注、表格、媒体和 OLE
对象等结构信号。它不能单独证明数学含义或页面渲染正确；完整要求见
[environment-and-validation.md](skills/word-math-doctor/references/environment-and-validation.md)。

默认输出只保存结构计数、部件名称和 SHA-256，不保存源文件路径或候选公式附近的正文片段。
上述 audits 目录已被 Git 忽略。仅在本地排查且绝不分享审计 JSON 时，才使用：

~~~bash
python skills/word-math-doctor/scripts/scan_docx_math.py manuscript.docx --include-snippets --include-paths --output audits/local-diagnostic.json
~~~

## 工作模式与验收

- 疑点扫描：只检查，不改原稿。
- 作者裁定：生成带建议正确写法和作者裁定列的 checklist。
- 裁定执行：只处理已确认项目，并扫描全文同类对象。
- 批注返修：逐项处理明确批注，保留歧义项。
- 全文规范化：完成转换、结构审计、渲染、重新保存和独立复核。

完整操作规则、格式规则、符号注册表、失败模式和验收清单位于
[SKILL.md](skills/word-math-doctor/SKILL.md)。

## GitHub、skills.sh 与 ChatGPT 的关系

- 发布到公开 GitHub 后，读者可以立即用本 README 中的 npx skills add 命令直接安装。
  直接安装不依赖 skills.sh 是否已有搜索结果。
- Skills CLI 的安装遥测可促使公开 Skill 自动出现在 skills.sh；但该过程没有人工提交命令，
  也不保证即时收录或一定能被搜索到。因此不要把 skills.sh 页面当作发布成功的唯一标志。
- GitHub 仓库公开不等于它会自动出现在 ChatGPT 网页端的插件目录。本仓库提供
  .codex-plugin/plugin.json 作为插件包元数据；若要面向 ChatGPT 网页端分发，仍需遵循
  该平台当时的插件安装或审核流程。

相关机制可参阅 [OpenAI Build skills](https://learn.chatgpt.com/docs/build-skills)、
[Skills CLI](https://github.com/antfu/skills-cli) 和
[Vercel 的 Skill 分享说明](https://vercel.com/kb/guide/agent-skills-creating-installing-and-sharing-reusable-agent-context)。

## 隐私与版本控制

本仓库不包含论文、数据或任何凭据。处理未公开文稿时，建议只在可控制的本地环境中执行，
不要将正文、批注、修订记录或文档元数据上传到第三方服务。每次返修应保留输入/输出文件、
SHA-256、结构审计结果和修改记录。

审计 JSON 的默认设置会避免保存原文片段和本地路径，但文件名、哈希、计数或经显式开启的
诊断信息仍可能识别文稿。不要将这些文件附在 Issue、Pull Request 或公开聊天记录中。

## 安全问题与引用

安全问题、隐私风险或可能导致文稿内容泄露的缺陷，请遵循
[SECURITY.md](SECURITY.md)，不要在公开 Issue 中附带真实文稿或审计输出。

如需在课程材料、论文或二次开发中引用本项目，请使用 [CITATION.cff](CITATION.cff) 中的
元数据，并保留 CC BY-NC 4.0 的署名与非商业使用条件。

## 贡献与本地验证

提交修改前运行：

~~~bash
python -m unittest discover -s tests -v
python skills/word-math-doctor/scripts/scan_docx_math.py --help
npx skills add . --list
~~~

贡献要求见 [CONTRIBUTING.md](CONTRIBUTING.md)。GitHub Actions 会在每次推送和 Pull Request
中执行相同的基本验证。

## 许可证

本项目采用 [CC BY-NC 4.0](LICENSE)。您可以在署名、保留许可链接并标明修改的前提下，
为非商业目的分享和改编本仓库内容。本仓库主要用于非商业教学与研究；商业课程、付费产品、
商业咨询、商业软件集成或边界不清的使用，须事先取得维护者许可。

CC BY-NC 4.0 的法律标准是非商业用途，并不等同于“仅限教学与研究”。具体边界以
[官方法律文本](https://creativecommons.org/licenses/by-nc/4.0/legalcode.en) 为准。
