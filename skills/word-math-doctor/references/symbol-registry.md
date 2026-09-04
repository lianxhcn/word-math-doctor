# 符号注册表与数学样式审计

## 1. 注册表字段

| 字段 | 说明 |
|---|---|
| `canonical` | 标准数学写法 |
| `category` | variable / function / operator / acronym / code-field |
| `structure` | 上下标、重音、花体、撇号和定界符 |
| `style` | italic / upright / text |
| `first_definition` | 首次定义位置 |
| `allowed_aliases` | 允许简写 |
| `forbidden_forms` | 禁止写法 |

## 2. 常见判定

- 单字母变量通常为数学斜体，例如 `b`、`B`、`i`、`α`。
- 下标索引通常为数学斜体，例如 `b_i`、`Q_{g,B}`。
- 说明性多字符下标按语义使用正体，例如 `min`、`max`、`coop`、`cal`、`day`、`year`。
- 函数名和运算符使用正体，例如 `exp`、`ln`、`BR`、`arg max`。
- MPC、HtM、NHtM、KKT 等缩写在正文中保持普通文本；在方程中再根据符号角色设置。
- 代码字段如 `excess_ARDays` 在正文叙述中保持正体；在方程中使用正体 OMML 文本。

## 3. 高风险结构

全文检索并比对以下表达式及其变体：

- `B^*` 与普通星号；
- `b_{-i}` 与 `b−i`、`b-ᵢ`；
- `\mathcal{M}` 与普通 `M`；
- `\widetilde{\Gamma}_{MB}`、`\widetilde{\Gamma}_{MS}` 与缩短或错误下标；
- `\bar{\chi}`、`\bar{p}`、`\bar{w}` 与 Unicode 组合上划线；
- `D_1`、`Q_{g,B}`、`\lambda_f`、`\mu_{E0}` 等复合下标；
- `BR_i(B\mathbf{1}_{N-1})`、`B^*=G(B^*)` 等固定点表达；
- 花体、黑体向量、波浪线、绝对值、范数、矩阵与条件评价竖线。

## 4. 语义核验顺序

1. 变量首次定义；
2. 当前公式；
3. 相邻正文解释；
4. 后续引用；
5. 正文与附录的多数写法；
6. 作者 checklist 裁定或 Word 批注。

上述来源仍不足以唯一确定表达式时，不修改；在待作者确认清单中列出候选写法和保留原因。
