---
name: word-math-doctor
description: 处理已有 Word 文档中的 Word 原生数学公式、OMML、Cambria Math、公式上下标与批注返修：扫描、裁定并在不改变数学含义和非目标格式的前提下修复可编辑公式。适用于全文或指定范围的公式规范化；不用于公式图片化、仅改字体或无依据改写数学含义。
license: CC-BY-NC-4.0
metadata:
  version: "0.2.0-beta"
  language: "zh-CN"
  short-description: "扫描、裁定并修复可编辑的 Word 原生数学公式。"
  author: "连享会"
  homepage: "https://github.com/lianxhcn/word-math-doctor"
  compatibility: "需要可读取 .docx 的本地环境；完整转换、渲染与重新保存应使用 Microsoft Word 或经验证的等价 OOXML/OMML 流程。"
---

# Word Math Doctor：原生公式校验、裁定与转换

## 1. 适用范围

用户提到下列任一情形时使用本 Skill：

- Word 原生数学公式、Word 自带数学公式、OMML 或 Cambria Math；
- 公式、变量、上下标、数学符号、公式编号或公式批注；
- 将普通文本、LaTeX、Unicode 伪公式或普通上下标转换为可编辑的 Word 原生公式；
- 对已存在的 OMML 公式做数学语义、正斜体、定界符和渲染复核；
- 根据 Word 批注或作者裁定统一返修公式。

不适用于以下任务：

- 仅把普通文字改成 Cambria Math；
- 将公式转成图片、SVG、EMF 或其他不可编辑对象；
- 没有公式或上下文依据时猜测、补写或改动数学含义；
- 将图中的公式 OCR 后静默替换。图中已有公式应单列为不可编辑来源，除非用户明确授权重建。

用户可显式调用：ChatGPT 使用 `@word-math-doctor`；Codex 使用 `$word-math-doctor`。

## 2. 开始前的能力与权限检查

本 Skill 是一套审计和返修工作流，不是 Word 加载项。开始前必须确认当前环境能够：

- 读取用户指定的 `.docx`，并将输出写入新文件；
- 直接检查 OOXML 中的 `m:oMath` 和 `m:oMathPara`，而不是只看屏幕字体；
- 对目标文档做实际渲染；全文规范化模式还应能用 Microsoft Word 或经验证的等价流程重新打开并保存；
- 计算文件哈希，并保存修改前后的结构审计结果。

若这些条件有任一项不具备，只能完成不依赖该能力的扫描或 checklist，且必须明确写出未完成的验收环节；不得声称已经完成 OMML 转换、渲染验收或全文规范化。具体环境要求见 [references/environment-and-validation.md](references/environment-and-validation.md)。

默认在用户可控制的本地环境中处理文稿。未经明确授权，不得将未公开论文、批注、修订痕迹或文档元数据上传到第三方服务。

## 3. 工作模式与完成边界

开始时必须判断本轮属于哪一种模式，并在回复中说明处理范围和不会处理的内容。

1. **疑点扫描**：只识别疑似数学表达式、符号层级、公式编排和环境风险，不修改原稿。
2. **作者裁定**：输出带「建议正确写法」和「作者裁定」列的 checklist；若 checklist 是 `.docx`，建议公式应为实际 OMML。Markdown checklist 只能作为语义预览，不能冒充 OMML。
3. **裁定执行**：只修改用户已确认的项目，并对全文同类对象做回归扫描。
4. **批注返修**：读取未解决批注，逐条处理或列入待确认项；只有用户已授权批量执行时，才能直接落实明确批注。
5. **全文规范化**：系统扫描、原位转换、结构审计、渲染、重新保存和独立复核。它需要用户明确授权处理的范围与数学语义权限。

完成 checklist 修订不等于完成全文 OMML 转换；存在 OMML 对象也不等于公式正确。只有全文扫描、结构审计、实际渲染、重新保存和独立复核均通过后，才能声明「全文数学公式规范化完成」。

## 4. 硬规则

- 行内公式使用真正的 `m:oMath`；独立公式使用 `m:oMathPara` 或专用公式段落中的完整 `m:oMath`。
- 禁止用 Cambria Math 普通文字、Word 普通上标/下标、Unicode 上下标或可见 LaTeX 冒充数学结构。
- 保留正确的原有 OMML，不做无意义重建；不得以图片替代原生公式。
- 原稿只读。输出必须是新文件，并记录输入与输出的 SHA-256、工作模式和处理范围。
- 默认不改变标题、正文样式、表格结构、图片、图表、脚注、页眉页脚、参考文献、批注和修订痕迹。公式转换引起的必要重排必须被单列报告，不得静默掩盖。
- 对照变量首次定义、相邻正文、编号公式、后续引用和全文多数写法。存在两种以上合理解释时，列入待作者确认清单。
- 不把代码、URL、DOI、文献中的 LaTeX 片段或非目标图形误判为待转换公式；它们应先分类，再决定是否属于任务范围。

详细格式规则见 [references/core-format-rules.md](references/core-format-rules.md)，符号判定见 [references/symbol-registry.md](references/symbol-registry.md)。

## 5. 标准工作流

### Step 1：确认底稿、范围和权限

明确最新底稿、目标章节或页码、工作模式，以及是否允许修正唯一可判定的数学错误。若用户没有明确全文授权，默认只扫描或处理已裁定项目。忽略以 `~$` 开头的 Word 临时文件。

### Step 2：建立可复核快照

记录文件名、SHA-256、段落、表格、图片、页眉页脚、脚注、尾注、批注、修订痕迹、`m:oMath`、`m:oMathPara`、普通上下标、Unicode 伪数学、可见 LaTeX 候选、OLE 与公式图片数量。优先运行随附的只读扫描器：

```bash
python scripts/scan_docx_math.py input.docx --output audits/input-audit.json
```

扫描器提供可比较的结构清单，不裁定数学语义，也不能代替逐页渲染。默认审计 JSON 不写入
源文件路径或候选公式附近的正文片段；只有在本地诊断且用户明确需要时，才附加
`--include-snippets --include-paths`。

### Step 3：建立符号注册表并分类扫描结果

从变量定义、独立公式、正文解释和附录建立符号注册表。对每个候选对象分类为：保留、局部修复、完整重建、非目标对象、待作者确认。批注和 checklist 只是入口，不是扫描边界；发现一个同类问题后，应扫描正文、附录、表格、图注、摘要和脚注中的同类对象。

### Step 4：生成可裁定清单

疑点清单至少包含位置、原稿显示或现象、疑点说明、建议写法、判断依据、置信程度和作者裁定。使用 [templates/decision-checklist-template.md](templates/decision-checklist-template.md)。若要向作者展示真正的 Word 原生公式，使用 `.docx` checklist；Markdown 中仅保留可读的线性或 LaTeX 预览，并标明其不是 OMML 对象。

### Step 5：原位修复与全文同类回归

只执行已授权或已裁定的项目。数学对象与中文说明分离：中文连接词留在普通正文 run，数学结构留在 OMML。表格采用原位最小修改，不重建整行或整表；同一单元格中的非数学文字、标点与换行应做前后快照比较。已有成熟公式编号与交叉引用默认保留。

### Step 6：结构、语义与版式验收

重新运行扫描器，并以输入快照为基线比较：

```bash
python scripts/scan_docx_math.py output.docx --baseline audits/input-audit.json --output audits/output-audit.json
```

核查 OMML 内部的正斜体、复合下标、星号、重音、花体、分式、矩阵、定界符、编号与交叉引用；再逐页渲染，重点检查所有修改页、长行内公式、复杂公式和表格。完整清单见 [references/acceptance-checklist.md](references/acceptance-checklist.md)。

### Step 7：重新打开保存与独立复核

全文规范化模式下，用 Microsoft Word 或已验证的等价流程重新打开并保存输出稿，再次渲染与审计。独立复核应由未参与主要修改的人或独立执行过程直接比对原稿、输出稿、审计记录和渲染图；若未完成，只能报告「主执行者复核通过，独立复核未完成」。

### Step 8：交付与如实声明

交付内容随模式调整：扫描模式至少交付疑点报告；执行模式至少交付新 `.docx`、修改记录、语义修正清单和待确认清单；全文规范化模式还应交付结构审计、渲染验收和独立复核报告。使用 [templates/modification-record-template.md](templates/modification-record-template.md)。

## 6. 失败条件

出现下列任一情况时，停止相关修改并输出异常清单：无法验证 OMML、存在未解决的数学歧义、非数学文字丢失、表格或版式发生未授权变化、公式裁切或编号失配、重新保存后结果变化，或没有完成所声明层级必需的验收。

已有公式图片、MathType/OLE 对象、可见 LaTeX 和普通上下标本身不是自动失败：先记录其位置和类别，再由用户决定保留、人工重建或排除出处理范围。不得把它们静默忽略后仍声称全文完成。

常见错误与处理边界见 [references/failure-modes.md](references/failure-modes.md)。
