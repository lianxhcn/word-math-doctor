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
import re
import sys
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}
ET.register_namespace("w", W)
ET.register_namespace("m", M)


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
        result: list[ET.Element] = []
        for item in spec:
            result.extend(math_items(item))
        return result
    if isinstance(spec, str):
        return [math_run(spec)]
    if not isinstance(spec, dict):
        raise PlanError("formula 节点必须是字符串或对象")
    kind = spec.get("kind")
    if kind == "text":
        value = spec.get("value")
        if not isinstance(value, str):
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
    properties.append(tabs)


def require_int(value: Any, name: str, minimum: int = 1) -> int:
    if not isinstance(value, int) or value < minimum:
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


def apply_operation(paragraphs: list[ET.Element], operation: dict[str, Any]) -> dict[str, Any]:
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
    if action == "inline":
        prefix = operation.get("prefix", "")
        suffix = operation.get("suffix", "")
        if not isinstance(prefix, str) or not isinstance(suffix, str):
            raise PlanError("inline 操作的 prefix 和 suffix 必须是字符串")
        if actual != prefix + operation.get("source_formula", "") + suffix:
            raise PlanError("inline 操作的 prefix/source_formula/suffix 未覆盖原段落文本")
        clear_contents_keep_properties(paragraph)
        if prefix:
            paragraph.append(word_run(prefix))
        paragraph.append(omml(formula))
        if suffix:
            paragraph.append(word_run(suffix))
    elif action == "display":
        number = normalize_number(operation.get("number"))
        if not actual.endswith(number):
            raise PlanError("display 的原段落必须以同一个半角公式编号结尾")
        center = require_int(operation.get("center_tab_twips", 4680), "center_tab_twips", 1)
        right = require_int(operation.get("right_tab_twips", 9360), "right_tab_twips", center + 1)
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
        paragraph.extend([tab_one, omml(formula), tab_two, word_run(number)])
    else:
        raise PlanError("action 只能是 inline 或 display")
    return {"paragraph": index, "action": action, "before": actual, "after": text_of_paragraph(paragraph)}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_plan(plan: Any) -> list[dict[str, Any]]:
    if not isinstance(plan, dict) or plan.get("schema") != "word-math-doctor/repair-plan/v1":
        raise PlanError("repair plan 的 schema 必须是 word-math-doctor/repair-plan/v1")
    operations = plan.get("operations")
    if not isinstance(operations, list) or not operations:
        raise PlanError("repair plan 必须含有非空 operations")
    if not all(isinstance(operation, dict) for operation in operations):
        raise PlanError("operations 的每一项必须是对象")
    positions = [operation.get("paragraph") for operation in operations]
    if len(set(positions)) != len(positions):
        raise PlanError("同一段落不得在一个 plan 中被修改两次")
    return operations


def repair(input_path: Path, plan_path: Path, output_path: Path, report_path: Path | None) -> dict[str, Any]:
    if input_path.resolve() == output_path.resolve():
        raise PlanError("输出文件必须与原文件不同")
    if output_path.exists():
        raise PlanError(f"输出文件已存在: {output_path}")
    try:
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise PlanError(f"repair plan 不是有效 JSON: {exc}") from exc
    operations = validate_plan(plan)
    input_hash_before = sha256(input_path)
    with zipfile.ZipFile(input_path) as archive:
        try:
            document_xml = archive.read("word/document.xml")
        except KeyError as exc:
            raise PlanError("输入文件不含 word/document.xml，不是可处理的 docx") from exc
        root = ET.fromstring(document_xml)
        paragraphs = list(root.iter(qn(W, "p")))
        before_omml = len(root.findall(".//m:oMath", NS))
        applied = [apply_operation(paragraphs, operation) for operation in operations]
        after_omml = len(root.findall(".//m:oMath", NS))
        serialized = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(output_path, "x", compression=zipfile.ZIP_DEFLATED) as output:
            for info in archive.infolist():
                data = serialized if info.filename == "word/document.xml" else archive.read(info.filename)
                copied_info = deepcopy(info)
                output.writestr(copied_info, data)
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
            "requires_word_render_check": any(item["action"] == "display" for item in applied),
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
    except (OSError, zipfile.BadZipFile, ET.ParseError, PlanError) as exc:
        print(f"repair_docx_math: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
