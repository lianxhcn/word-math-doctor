"""Create a separate academic-font copy of the original fictional test DOCX.

This fixture preparation is deliberately separate from formula conversion.
Requires python-docx, only for producing the manual Word compatibility fixture.
The actual repair script has no third-party dependencies.
"""
from pathlib import Path
import argparse
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


def prepare(source: Path, output: Path) -> None:
    # 保留测试材料的正文与公式类型，仅准备明确的学术字体基准副本。
    if source.resolve() == output.resolve() or output.exists():
        raise ValueError("测试副本必须使用新的输出路径")
    doc = Document(source)
    for name in ("Normal", "Title", "Heading 1"):
        style = doc.styles[name]
        style.font.name = "Times New Roman"
        style.font.size = Pt(12 if name == "Normal" else 16)
        style.font.italic = False
        fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
        fonts.set(qn("w:eastAsia"), "宋体")
        for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn("w:" + key), None)
    # 直接格式覆盖只作用于测试副本的普通文字，不触碰 m:oMath。
    for run in doc.element.body.iter(qn("w:r")):
        props = run.get_or_add_rPr()
        fonts = props.get_or_add_rFonts()
        fonts.set(qn("w:ascii"), "Times New Roman")
        fonts.set(qn("w:hAnsi"), "Times New Roman")
        fonts.set(qn("w:eastAsia"), "宋体")
        for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn("w:" + key), None)
        props.get_or_add_sz().val = Pt(12)
        props.get_or_add_i().val = False
        props.get_or_add_iCs().val = False
    # 这段说明是测试材料的任务说明，避免「只读审计」与本次授权转换冲突。
    doc.paragraphs[1].text = "这是一份虚构测试材料。仅转换明确授权的公式，另存修订稿；原文件、化学式和非目标内容保持不变。"
    doc.save(output)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="生成独立的学术字体测试副本")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    prepare(args.source, args.output)
