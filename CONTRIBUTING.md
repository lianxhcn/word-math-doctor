# 贡献说明

欢迎提交文档修订、测试、扫描器改进和可复现的缺陷报告。

## 提交前检查

1. 不要提交论文原稿、批注、作者身份信息、教学材料、学生作业、数据、审计 JSON 或访问凭据。
2. 保持 Skill 的安全边界：不能把推测的数学含义写回文稿，不能将公式图片静默替换为 OMML。
3. 运行以下检查：

~~~bash
python -m unittest discover -s tests -v
python skills/word-math-doctor/scripts/scan_docx_math.py --help
~~~

4. 对影响工作流的改动，请在 README 或 Changelog 中说明用户可观察到的行为变化。

## Issues 与 Pull Requests

- 缺陷报告请提供最小化、脱敏后的 OOXML 结构或虚构样例，不要上传真实未公开文稿。
- 请说明预期结果、实际结果、操作系统、Python 版本和 Word 版本（如适用）。
- 提交贡献即表示您有权提交该内容，并同意该内容在 CC BY-NC 4.0 下发布。
