#!/usr/bin/env python3
"""Apply a reviewed, bounded OMML repair plan to a .docx file.

This tool deliberately has no automatic formula detection or inference.  A
human/agent must first review a plan that names each target paragraph, its
expected original text, and the desired mathematical structure.  The script
then makes only those changes, writes a separate output file, and records a
machine-readable report.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile
import zipfile
from copy import deepcopy
from xml.parsers import expat
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W, "m": M}
ET.register_namespace("w", W)
ET.register_namespace("m", M)
ET.register_namespace("mc", MC)


class PlanError(ValueError):
    """The repair plan is incomplete, unsafe, or does not match the source."""


def qn(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def element(namespace: str, name: str, **attributes: str) -> ET.Element:
    node = ET.Element(qn(namespace, name))
    for key, value in attributes.items():
        node.set(qn(namespace, key), str(value))
    return node


def text_of_paragraph(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qn(W, "t")))


def set_text(run: ET.Element, value: str) -> None:
    text = element(W, "t")
    if value.startswith(" ") or value.endswith(" "):
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = value
    run.append(text)


def word_run(value: str) -> ET.Element:
    run = element(W, "r")
    set_text(run, value)
    return run


def math_run(value: str) -> ET.Element:
    run = element(M, "r")
    text = element(M, "t")
    if value.startswith(" ") or value.endswith(" "):
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = value
    run.append(text)
    return run


def math_argument(items: list[ET.Element]) -> ET.Element:
    argument = element(M, "e")
    for item in items:
        argument.append(item)
    return argument


def math_slot(name: str, items: list[ET.Element]) -> ET.Element:
    slot = element(M, name)
    for item in items:
        slot.append(item)
    return slot


def math_items(spec: Any) -> list[ET.Element]:
    """Compile a small, explicit JSON formula AST into OMML children."""
    if isinstance(spec, list):
        if not spec:
            raise PlanError("公式结构不得为空")
        result: list[ET.Element] = []
        for item in spec:
            result.extend(math_items(item))
        return result
    if isinstance(spec, str):
        if not spec:
            raise PlanError("公式文本节点不得为空")
        return [math_run(spec)]
    if not isinstance(spec, dict):
        raise PlanError("formula 节点必须是字符串或对象")
    kind = spec.get("kind")
    if kind == "text":
        value = spec.get("value")
        if not isinstance(value, str) or not value:
            raise PlanError("text 节点需要字符串 value")
        return [math_run(value)]
    if kind == "sequence":
        items = spec.get("items")
        if not isinstance(items, list) or not items:
            raise PlanError("sequence 节点需要非空 items")
        result: list[ET.Element] = []
        for item in items:
            result.extend(math_items(item))
        return result
    if kind in {"sup", "sub"}:
        base = math_items(spec.get("base"))
        value = math_items(spec.get("exponent" if kind == "sup" else "subscript"))
        node = element(M, "sSup" if kind == "sup" else "sSub")
        node.append(math_slot("e", base))
        node.append(math_slot("sup" if kind == "sup" else "sub", value))
        return [node]
    if kind == "fraction":
        numerator = math_items(spec.get("numerator"))
        denominator = math_items(spec.get("denominator"))
        node = element(M, "f")
        node.append(math_slot("num", numerator))
        node.append(math_slot("den", denominator))
        return [node]
    if kind == "integral":
        lower = math_items(spec.get("lower"))
        upper = math_items(spec.get("upper"))
        body = math_items(spec.get("body"))
        node = element(M, "nary")
        properties = element(M, "naryPr")
        properties.append(element(M, "chr", val="∫"))
        properties.append(element(M, "limLoc", val="subSup"))
        node.append(properties)
        sub = element(M, "sub")
        for item in lower:
            sub.append(item)
        sup = element(M, "sup")
        for item in upper:
            sup.append(item)
        node.extend([sub, sup, math_argument(body)])
        return [node]
    raise PlanError(f"不支持的 formula kind: {kind!r}")


def omml(spec: Any) -> ET.Element:
    node = element(M, "oMath")
    for item in math_items(spec):
        node.append(item)
    return node


def paragraph_properties(paragraph: ET.Element) -> ET.Element | None:
    return paragraph.find(qn(W, "pPr"))


def clear_contents_keep_properties(paragraph: ET.Element) -> None:
    properties = paragraph_properties(paragraph)
    for child in list(paragraph):
        if child is not properties:
            paragraph.remove(child)


def add_display_tabs(properties: ET.Element, center: int, right: int) -> None:
    for tabs in list(properties.findall(qn(W, "tabs"))):
        properties.remove(tabs)
    tabs = element(W, "tabs")
    tabs.append(element(W, "tab", val="center", pos=str(center)))
    tabs.append(element(W, "tab", val="right", pos=str(right)))
    # CT_PPr has an ordered content model: tabs precedes spacing/ind/jc/rPr.
    following = {"suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
                 "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
                 "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
                 "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap",
                 "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"}
    position = next((i for i, child in enumerate(properties)
                     if child.tag in {qn(W, name) for name in following}), len(properties))
    properties.insert(position, tabs)


def require_int(value: Any, name: str, minimum: int = 1) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise PlanError(f"{name} 必须是不小于 {minimum} 的整数")
    return value


def normalize_number(value: Any) -> str:
    if isinstance(value, int):
        value = str(value)
    if not isinstance(value, str):
        raise PlanError("display 操作需要字符串或整数 number")
    match = re.fullmatch(r"(?:\(([0-9]+)\)|([0-9]+))", value)
    if not match:
        raise PlanError("公式编号只能是半角数字或半角括号包围的数字，例如 1 或 (1)")
    return f"({match.group(1) or match.group(2)})"


def simple_runs(paragraph: ET.Element) -> list[ET.Element]:
    """Fail closed on fields, anchors, revisions and nested/non-text content."""
    runs = []
    for child in paragraph:
        if child.tag == qn(W, "pPr"):
            if child.find(".//w:pPrChange", NS) is not None:
                raise PlanError("目标段落含属性修订，须由 Word/人工处理")
            continue
        if child.tag != qn(W, "r"):
            raise PlanError("目标段落含公式、书签、批注、域、超链接或修订，不能安全替换")
        if any(node.tag not in {qn(W, "rPr"), qn(W, "t")} for node in child):
            raise PlanError("目标 run 含非纯文本内容，不能安全替换")
        if len(child.findall("w:t", NS)) > 1:
            raise PlanError("目标 run 含多个文本节点，须先人工检查")
        text = child.find("w:t", NS)
        if text is None or not text.text:
            raise PlanError("目标段落含空 run，无法可靠保存其边界格式")
        if child.find(".//w:rPrChange", NS) is not None:
            raise PlanError("目标 run 含字符属性修订")
        runs.append(child)
    if any(not isinstance(node.tag, str) or node.tag.startswith("{" + MC + "}") or
           any(key.startswith("{" + MC + "}") for key in node.attrib)
           for node in paragraph.iter()):
        raise PlanError("目标段落含兼容性选择/前缀值，不能安全重写")
    return runs


def slice_runs(runs: list[ET.Element], start: int, end: int) -> list[ET.Element]:
    """Copy the selected text range, retaining source run/t attributes and rPr."""
    result = []
    offset = 0
    for run in runs:
        node = run.find("w:t", NS)
        value = node.text or "" if node is not None else ""
        stop = offset + len(value)
        if start <= offset and stop <= end and value:
            result.append(deepcopy(run))
        elif offset < end and stop > start:
            copied = deepcopy(run)
            text = copied.find("w:t", NS)
            text.text = value[max(0, start - offset):min(len(value), end - offset)]
            if text.text.startswith(" ") or text.text.endswith(" "):
                text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            result.append(copied)
        offset = stop
    return result


def paragraph_property_layers(paragraph: ET.Element, styles: ET.Element | None) -> list[ET.Element]:
    """Inspect defaults and the paragraph's basedOn chain, then direct pPr."""
    direct = paragraph_properties(paragraph)
    layers = []
    if styles is not None:
        default = styles.find("w:docDefaults/w:pPrDefault/w:pPr", NS)
        if default is not None:
            layers.append(default)
        style_id = direct.find("w:pStyle", NS).get(qn(W, "val")) if direct is not None and direct.find("w:pStyle", NS) is not None else None
        paragraph_styles = {s.get(qn(W, "styleId")): s for s in styles.findall("w:style", NS)
                            if s.get(qn(W, "type")) == "paragraph"}
        if style_id is None:
            style_id = next((k for k, s in paragraph_styles.items()
                             if s.get(qn(W, "default")) in {"1", "true", "on"}), None)
        chain = []
        seen = set()
        while style_id:
            if style_id in seen or style_id not in paragraph_styles:
                raise PlanError("段落样式继承链缺失或循环，不能计算版心")
            seen.add(style_id)
            style = paragraph_styles[style_id]
            props = style.find("w:pPr", NS)
            if props is not None:
                chain.append(props)
            parent = style.find("w:basedOn", NS)
            style_id = parent.get(qn(W, "val")) if parent is not None else None
        layers.extend(reversed(chain))
    elif direct is not None and direct.find("w:pStyle", NS) is not None:
        raise PlanError("缺少 styles.xml，无法解析段落样式")
    if direct is not None:
        layers.append(direct)
    return layers


def display_layout(root: ET.Element, paragraph: ET.Element, styles: ET.Element | None,
                   settings: ET.Element | None) -> tuple[int, int]:
    """Support explicit section geometry in unindented, horizontal body text."""
    body = root.find("w:body", NS)
    if body is None or paragraph not in list(body):
        raise PlanError("独立公式不在普通正文中(表格/文本框等)，须人工排版")
    parents = {child: parent for parent in root.iter() for child in parent}
    for section in root.iter(qn(W, "sectPr")):
        parent = parents.get(section)
        if parent is body:
            continue
        if parent is None or parent.tag != qn(W, "pPr") or parents.get(parent) not in list(body):
            raise PlanError("文档含嵌套容器中的分节属性，不能可靠定位所在节")
    if settings is not None and any(settings.find("w:" + name, NS) is not None
                                    for name in ("mirrorMargins", "gutterAtTop")):
        raise PlanError("镜像页边距/装订线设置暂不支持")
    for props in paragraph_property_layers(paragraph, styles):
        for node in props:
            local = node.tag.split("}")[-1]
            if local in {"numPr", "framePr", "sectPr", "pPrChange"}:
                raise PlanError("目标段落含自动编号、框架、分节或修订属性")
            if local == "ind" and any(value != "0" for value in node.attrib.values()):
                raise PlanError("非零段落缩进暂不支持；不得按整页宽度计算制表位")
            if local == "tabs" and len(node):
                raise PlanError("段落/样式已有自定义制表位，须人工裁定")
            if local == "jc" and node.get(qn(W, "val")) not in {"left", "start"}:
                raise PlanError("非左对齐段落暂不支持")
            if local in {"bidi", "mirrorIndents", "adjustRightInd"} and node.get(qn(W, "val"), "1") not in {"0", "false", "off"}:
                raise PlanError("双向文字/镜像缩进暂不支持")
            if local == "textDirection" and node.get(qn(W, "val")) != "lrTb":
                raise PlanError("非水平文字暂不支持")
    # Section properties close the section; do not use the preceding section.
    section = None
    for child in list(body)[list(body).index(paragraph):]:
        section = child if child.tag == qn(W, "sectPr") else child.find("w:pPr/w:sectPr", NS)
        if section is not None:
            break
    if section is None:
        raise PlanError("无法定位公式所在节")
    if section.find("w:sectPrChange", NS) is not None:
        raise PlanError("所在节含修订属性")
    cols = section.find("w:cols", NS)
    if cols is not None and (cols.get(qn(W, "num"), "1") != "1" or len(cols)):
        raise PlanError("分栏版式暂不支持")
    if section.find("w:textDirection", NS) is not None or section.find("w:bidi", NS) is not None:
        raise PlanError("节文字方向暂不支持")
    kind = section.find("w:type", NS)
    if kind is not None and kind.get(qn(W, "val")) in {"continuous", "nextColumn"}:
        raise PlanError("连续/分栏分节暂不支持")
    size, margins = section.find("w:pgSz", NS), section.find("w:pgMar", NS)
    if size is None or margins is None:
        raise PlanError("所在节缺少显式页面宽度或页边距，禁止猜测")
    try:
        width = int(size.get(qn(W, "w")))
        left = int(margins.get(qn(W, "left")))
        right = int(margins.get(qn(W, "right")))
        gutter = int(margins.get(qn(W, "gutter"), "0"))
    except (TypeError, ValueError) as exc:
        raise PlanError("页面宽度/页边距不是有效 twips") from exc
    if min(left, right) < 0 or gutter != 0 or width - left - right <= 0:
        raise PlanError("页面几何无效或有装订线，须人工排版")
    usable = width - left - right
    return usable // 2, usable


def number_runs(runs: list[ET.Element], start: int, end: int, style: Any) -> list[ET.Element]:
    result = slice_runs(runs, start, end)
    if not result:
        raise PlanError("无法读取原公式编号格式")
    if style is not None and (not isinstance(style, dict) or set(style) - {"font", "size_half_points"}):
        raise PlanError("number_style 只支持 font 与 size_half_points")
    for run in result:
        props = run.find("w:rPr", NS)
        if props is None:
            props = element(W, "rPr")
            run.insert(0, props)
        # Preserve source font/size/character style; explicitly make labels upright.
        for name in ("i", "iCs"):
            existing = props.find("w:" + name, NS)
            if existing is not None:
                existing.set(qn(W, "val"), "0")
            else:
                node = element(W, name, val="0")
                earlier = {"rStyle", "rFonts", "b", "bCs"} | ({"i"} if name == "iCs" else set())
                pos = 0
                while pos < len(props) and props[pos].tag in {qn(W, key) for key in earlier}:
                    pos += 1
                props.insert(pos, node)
        if style and "font" in style:
            font = style["font"]
            if not isinstance(font, str) or not font.strip():
                raise PlanError("number_style.font 必须是非空字体名")
            fonts = props.find("w:rFonts", NS)
            if fonts is None:
                fonts = element(W, "rFonts")
                props.insert(1 if props.find("w:rStyle", NS) is not None else 0, fonts)
            for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                fonts.attrib.pop(qn(W, key), None)
            for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                fonts.set(qn(W, key), font)
        if style and "size_half_points" in style:
            value = str(require_int(style["size_half_points"], "size_half_points"))
            for name in ("sz", "szCs"):
                existing = props.find("w:" + name, NS)
                if existing is not None:
                    existing.set(qn(W, "val"), value)
                else:
                    node = element(W, name, val=value)
                    following = {"szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"}
                    if name == "szCs":
                        following.discard("szCs")
                    pos = next((i for i, child in enumerate(props) if child.tag in {qn(W, key) for key in following}), len(props))
                    props.insert(pos, node)
    return result


def apply_operation(paragraphs: list[ET.Element], operation: dict[str, Any],
                    root: ET.Element, styles: ET.Element | None,
                    settings: ET.Element | None) -> dict[str, Any]:
    index = require_int(operation.get("paragraph"), "paragraph", 1)
    if index > len(paragraphs):
        raise PlanError(f"paragraph={index} 超出文档段落数量 {len(paragraphs)}")
    paragraph = paragraphs[index - 1]
    expected = operation.get("expected_text")
    actual = text_of_paragraph(paragraph)
    if not isinstance(expected, str) or actual != expected:
        raise PlanError(
            f"第 {index} 段的 expected_text 不匹配；为避免误改已停止。"
        )
    action = operation.get("action")
    formula = operation.get("formula")
    runs = simple_runs(paragraph)
    math = omml(formula)
    details = {}
    if action == "inline":
        prefix = operation.get("prefix", "")
        suffix = operation.get("suffix", "")
        if not isinstance(prefix, str) or not isinstance(suffix, str):
            raise PlanError("inline 操作的 prefix 和 suffix 必须是字符串")
        source_formula = operation.get("source_formula")
        if not isinstance(source_formula, str) or not source_formula or actual != prefix + source_formula + suffix:
            raise PlanError("inline 操作的 prefix/source_formula/suffix 未覆盖原段落文本")
        before = slice_runs(runs, 0, len(prefix))
        after = slice_runs(runs, len(prefix) + len(source_formula), len(actual))
        clear_contents_keep_properties(paragraph)
        paragraph.extend(before + [math] + after)
    elif action == "display":
        number = normalize_number(operation.get("number"))
        if not actual.endswith(number):
            raise PlanError("display 的原段落必须以同一个半角公式编号结尾")
        source_formula = operation.get("source_formula")
        if not isinstance(source_formula, str) or not source_formula.strip():
            raise PlanError("display 必须明确声明 source_formula，不得隐式删除整段文字")
        remainder = actual[len(source_formula): -len(number)]
        if not actual.startswith(source_formula) or not remainder or not remainder.isspace():
            raise PlanError("display 必须仅含 source_formula、间隔空白和编号；解释文字不得删除")
        center, right = display_layout(root, paragraph, styles, settings)
        for name, calculated in (("center_tab_twips", center), ("right_tab_twips", right)):
            if name in operation and require_int(operation[name], name) != calculated:
                raise PlanError(f"{name} 与实际版心不符；禁止越界或偏心制表位")
        labels = number_runs(runs, len(actual) - len(number), len(actual), operation.get("number_style"))
        details = {"center_tab_twips": center, "right_tab_twips": right,
                   "number_style_source": "explicit_override" if operation.get("number_style") else "original_run_and_document_style"}
        clear_contents_keep_properties(paragraph)
        properties = paragraph_properties(paragraph)
        if properties is None:
            properties = element(W, "pPr")
            paragraph.insert(0, properties)
        add_display_tabs(properties, center, right)
        tab_one = element(W, "r")
        tab_one.append(element(W, "tab"))
        tab_two = element(W, "r")
        tab_two.append(element(W, "tab"))
        paragraph.extend([tab_one, math, tab_two] + labels)
    else:
        raise PlanError("action 只能是 inline 或 display")
    return {"paragraph": index, "action": action, "before": actual,
            "after_word_text": text_of_paragraph(paragraph),
            "formula_structure": deepcopy(formula),
            **details}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_spans(xml: bytes) -> list[tuple[int, int]]:
    """Find exact UTF-8 byte ranges with Expat, without regex-parsing XML.

    Only selected paragraphs are serialized. Root namespace declarations,
    compatibility attributes and every byte outside those ranges are retained.
    """
    if b"<!DOCTYPE" in xml or b"<!ENTITY" in xml:
        raise PlanError("不支持 DTD/实体声明")
    declaration = re.match(rb"\s*<\?xml[^?]*encoding=['\"]([^'\"]+)", xml)
    if declaration and declaration.group(1).lower() not in {b"utf-8", b"utf8"}:
        raise PlanError("仅支持 UTF-8 document.xml")
    parser = expat.ParserCreate(namespace_separator="}")
    spans = []
    stack = []

    def tag_end(offset: int) -> int:
        quote = None
        for index in range(offset, len(xml)):
            char = xml[index]
            if quote:
                if char == quote:
                    quote = None
            elif char in (34, 39):
                quote = char
            elif char == 62:
                return index + 1
        raise PlanError("XML 标签未闭合")

    def start(name: str, attributes: dict) -> None:
        if name == W + "}p":
            offset = parser.CurrentByteIndex
            # Empty paragraphs end at the start of the next token.
            opening_end = tag_end(offset)
            empty = xml[offset:opening_end].rstrip().endswith(b"/>")
            stack.append((len(spans), offset, opening_end if empty else None))
            spans.append(None)

    def end(name: str) -> None:
        if name == W + "}p":
            index, offset, empty_end = stack.pop()
            stop = empty_end if empty_end is not None else tag_end(parser.CurrentByteIndex)
            spans[index] = (offset, stop)

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(xml, True)
    return spans


def patch_document(xml: bytes, paragraphs: list[ET.Element], operations: list[dict]) -> bytes:
    spans = paragraph_spans(xml)
    if len(spans) != len(paragraphs):
        raise PlanError("XML 字节定位与解析结果不一致")
    patches = []
    for operation in operations:
        index = operation["paragraph"] - 1
        begin, end = spans[index]
        paragraph = deepcopy(paragraphs[index])
        paragraph.tail = None
        replacement = ET.tostring(paragraph, encoding="utf-8")
        patches.append((begin, end, replacement))
    patches.sort()
    if any(left[1] > right[0] for left, right in zip(patches, patches[1:])):
        raise PlanError("目标段落范围嵌套或重叠")
    pieces = []
    offset = 0
    for begin, end, replacement in patches:
        pieces.extend([xml[offset:begin], replacement])
        offset = end
    pieces.append(xml[offset:])
    result = b"".join(pieces)
    ET.fromstring(result)
    return result


def protected_paragraphs(root: ET.Element) -> set[ET.Element]:
    """Reject targets inside a range even when anchors sit in other paragraphs."""
    protected = set()
    active = set()
    field_depth = 0
    for node in root.iter():
        if node.tag == qn(W, "p") and (active or field_depth):
            protected.add(node)
        if node.tag in {qn(W, "bookmarkStart"), qn(W, "commentRangeStart")}:
            active.add((node.tag, node.get(qn(W, "id"))))
        elif node.tag in {qn(W, "bookmarkEnd"), qn(W, "commentRangeEnd")}:
            start_tag = node.tag.replace("End", "Start")
            active.discard((start_tag, node.get(qn(W, "id"))))
        elif node.tag == qn(W, "fldChar"):
            kind = node.get(qn(W, "fldCharType"))
            if kind == "begin":
                field_depth += 1
            elif kind == "end":
                field_depth = max(0, field_depth - 1)
    return protected


def validate_plan(plan: Any) -> list[dict[str, Any]]:
    if not isinstance(plan, dict) or plan.get("schema") != "word-math-doctor/repair-plan/v1":
        raise PlanError("repair plan 的 schema 必须是 word-math-doctor/repair-plan/v1")
    operations = plan.get("operations")
    if not isinstance(operations, list) or not operations:
        raise PlanError("repair plan 必须含有非空 operations")
    if not all(isinstance(operation, dict) for operation in operations):
        raise PlanError("operations 的每一项必须是对象")
    positions = [require_int(operation.get("paragraph"), "paragraph") for operation in operations]
    if len(set(positions)) != len(positions):
        raise PlanError("同一段落不得在一个 plan 中被修改两次")
    return operations


def repair(input_path: Path, plan_path: Path, output_path: Path, report_path: Path | None) -> dict[str, Any]:
    if input_path.suffix.lower() != ".docx" or output_path.suffix.lower() != ".docx":
        raise PlanError("限定修复仅支持 .docx；.docm 只可审计")
    paths = [input_path, plan_path, output_path] + ([report_path] if report_path else [])
    if len({path.resolve() for path in paths}) != len(paths):
        raise PlanError("原稿、计划、输出及报告路径必须互不相同")
    if output_path.exists():
        raise PlanError(f"输出文件已存在: {output_path}")
    if report_path and report_path.exists():
        raise PlanError("报告文件已存在，禁止覆盖")
    try:
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise PlanError(f"repair plan 不是有效 JSON: {exc}") from exc
    operations = validate_plan(plan)
    input_hash_before = sha256(input_path)
    with zipfile.ZipFile(input_path) as archive:
        if len(set(archive.namelist())) != len(archive.namelist()):
            raise PlanError("DOCX ZIP 含重复部件")
        try:
            document_xml = archive.read("word/document.xml")
        except KeyError as exc:
            raise PlanError("输入文件不含 word/document.xml，不是可处理的 docx") from exc
        root = ET.fromstring(document_xml, parser=ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True)))
        styles = ET.fromstring(archive.read("word/styles.xml")) if "word/styles.xml" in archive.namelist() else None
        settings = ET.fromstring(archive.read("word/settings.xml")) if "word/settings.xml" in archive.namelist() else None
        paragraphs = list(root.iter(qn(W, "p")))
        protected = protected_paragraphs(root)
        body = root.find("w:body", NS)
        for operation in operations:
            index = operation["paragraph"] - 1
            if index >= len(paragraphs):
                raise PlanError("paragraph 超出文档段落数量")
            target = paragraphs[index]
            if body is None or target not in list(body) or target in protected:
                raise PlanError("目标位于表格/嵌套结构或跨段书签、批注、域范围内，须人工处理")
        before_omml = len(root.findall(".//m:oMath", NS))
        applied = [apply_operation(paragraphs, operation, root, styles, settings) for operation in operations]
        after_omml = len(root.findall(".//m:oMath", NS))
        serialized = patch_document(document_xml, paragraphs, operations)
        # Final output appears only after a complete ZIP has been written.
        # link() provides no-clobber publication even if another process creates
        # the requested output while this script is running.
        with tempfile.NamedTemporaryFile(dir=output_path.parent, prefix=".wmd-", suffix=".docx", delete=False) as temporary:
            temporary_path = Path(temporary.name)
        try:
            with zipfile.ZipFile(temporary_path, "w", compression=zipfile.ZIP_DEFLATED) as output:
                output.comment = archive.comment
                for info in archive.infolist():
                    data = serialized if info.filename == "word/document.xml" else archive.read(info.filename)
                    output.writestr(deepcopy(info), data)
            if sha256(input_path) != input_hash_before:
                raise PlanError("转换期间原稿发生变化，已停止")
            os.link(temporary_path, output_path)
        finally:
            temporary_path.unlink(missing_ok=True)
    report = {
        "schema": "word-math-doctor/repair-report/v1",
        "input_sha256": input_hash_before,
        "output_sha256": sha256(output_path),
        "operations_applied": applied,
        "omml_before": before_omml,
        "omml_after": after_omml,
        "untouched_parts_copied": True,
        "validation": {
            "input_unchanged": sha256(input_path) == input_hash_before,
            "non_target_document_bytes_preserved": True,
            "requires_word_render_check": True,
            "word_open_save_verified": False,
            "limitations": ["未验证公式实际宽度、编号碰撞、字体替代、裁切或跨页；必须在 Word 中打开、编辑、保存并复核。"],
        },
    }
    if report_path:
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="按已审核的计划限定修复 Word 数学公式")
    parser.add_argument("input", type=Path, help="原始 .docx")
    parser.add_argument("plan", type=Path, help="已审核的 repair plan JSON")
    parser.add_argument("output", type=Path, help="新的 .docx 输出路径，必须不存在")
    parser.add_argument("--report", type=Path, help="可选的 JSON 修改报告")
    args = parser.parse_args()
    try:
        report = repair(args.input, args.plan, args.output, args.report)
    except (OSError, zipfile.BadZipFile, ET.ParseError, expat.ExpatError, PlanError) as exc:
        print(f"repair_docx_math: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
