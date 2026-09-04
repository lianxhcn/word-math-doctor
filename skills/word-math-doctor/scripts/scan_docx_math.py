#!/usr/bin/env python3
"""Read-only structural inventory for Word mathematical content.

The scanner does not modify a document and cannot decide mathematical meaning.
It is designed to create reproducible before/after OOXML inventories for
Word Math Doctor.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


TOOL_VERSION = "0.2.0-beta"

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

CONTENT_PART = re.compile(
    r"^word/(?:document|footnotes|endnotes|comments|header\d+|footer\d+)\.xml$"
)
UNICODE_SUPSUB = re.compile("[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿᶦ₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ]")
LATEX_CANDIDATE = re.compile(r"(?<!\\)\\(?:[A-Za-z]+|[{}_^])")
TEXT_MATH_CANDIDATE = re.compile(
    r"(?<![A-Za-z0-9])(?:[A-Za-zα-ωΑ-Ω]\s*(?:[_^]\s*\{?[A-Za-z0-9+\-]+\}?|[=<>]\s*[A-Za-z0-9α-ωΑ-Ω]))"
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def count(root: ET.Element, xpath: str) -> int:
    return len(root.findall(xpath, NS))


def candidate_snippets(text: str, pattern: re.Pattern[str], limit: int = 5) -> list[str]:
    snippets: list[str] = []
    for match in pattern.finditer(text):
        start = max(0, match.start() - 30)
        end = min(len(text), match.end() + 30)
        snippet = re.sub(r"\s+", " ", text[start:end]).strip()
        if snippet and snippet not in snippets:
            snippets.append(snippet)
        if len(snippets) >= limit:
            break
    return snippets


def inspect_part(
    name: str, raw_xml: bytes, *, include_snippets: bool
) -> dict[str, Any]:
    root = ET.fromstring(raw_xml)
    text = "".join((node.text or "") for node in root.findall(".//w:t", NS))
    text += "".join((node.text or "") for node in root.findall(".//w:instrText", NS))

    latex_matches = LATEX_CANDIDATE.findall(text)
    math_matches = TEXT_MATH_CANDIDATE.findall(text)

    inventory = {
        "part": name,
        "paragraphs": count(root, ".//w:p"),
        "tables": count(root, ".//w:tbl"),
        "table_cells": count(root, ".//w:tc"),
        "omml_objects": count(root, ".//m:oMath"),
        "omml_paragraphs": count(root, ".//m:oMathPara"),
        "legacy_superscript_runs": len(
            root.findall(".//w:vertAlign[@w:val='superscript']", NS)
        ),
        "legacy_subscript_runs": len(
            root.findall(".//w:vertAlign[@w:val='subscript']", NS)
        ),
        "unicode_superscript_or_subscript_characters": len(UNICODE_SUPSUB.findall(text)),
        "visible_latex_candidates": len(latex_matches),
        "text_math_candidates": len(math_matches),
        "comment_ranges": count(root, ".//w:commentRangeStart"),
        "comments": count(root, ".//w:comment"),
        "tracked_insertions": count(root, ".//w:ins"),
        "tracked_deletions": count(root, ".//w:del"),
        "drawings": count(root, ".//w:drawing"),
        "vml_pictures": count(root, ".//w:pict"),
        "ole_references": count(root, ".//o:OLEObject"),
    }
    if include_snippets:
        inventory["text_math_candidate_snippets"] = candidate_snippets(
            text, TEXT_MATH_CANDIDATE
        )
        inventory["latex_candidate_snippets"] = candidate_snippets(
            text, LATEX_CANDIDATE
        )
    return inventory


COUNT_FIELDS = (
    "paragraphs",
    "tables",
    "table_cells",
    "omml_objects",
    "omml_paragraphs",
    "legacy_superscript_runs",
    "legacy_subscript_runs",
    "unicode_superscript_or_subscript_characters",
    "visible_latex_candidates",
    "text_math_candidates",
    "comment_ranges",
    "comments",
    "tracked_insertions",
    "tracked_deletions",
    "drawings",
    "vml_pictures",
    "ole_references",
)


def make_audit(
    path: Path, *, include_snippets: bool = False, include_paths: bool = False
) -> dict[str, Any]:
    if not path.is_file():
        raise ValueError(f"File does not exist: {path}")
    if path.suffix.lower() not in {".docx", ".docm"}:
        raise ValueError("Expected a .docx or .docm file")

    try:
        archive = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise ValueError("The input is not a readable Office Open XML ZIP package") from exc

    with archive:
        names = archive.namelist()
        content_parts = sorted(name for name in names if CONTENT_PART.match(name))
        media_files = sorted(name for name in names if name.startswith("word/media/"))
        embedded_objects = sorted(
            name
            for name in names
            if name.startswith("word/embeddings/") or name.startswith("word/activeX/")
        )

        inspected_parts: list[dict[str, Any]] = []
        parse_errors: list[dict[str, str]] = []
        for name in content_parts:
            try:
                inspected_parts.append(
                    inspect_part(
                        name, archive.read(name), include_snippets=include_snippets
                    )
                )
            except (ET.ParseError, KeyError, UnicodeDecodeError) as exc:
                parse_errors.append({"part": name, "error": str(exc)})

    counts = {field: sum(part[field] for part in inspected_parts) for field in COUNT_FIELDS}
    counts["media_files"] = len(media_files)
    counts["embedded_object_files"] = len(embedded_objects)
    counts["unparsed_content_parts"] = len(parse_errors)

    source = {"sha256": sha256_file(path)}
    if include_paths:
        source["path"] = str(path)

    return {
        "schema": "word-math-doctor/docx-math-audit/v2",
        "tool_version": TOOL_VERSION,
        "source": source,
        "privacy": {
            "source_paths": "included" if include_paths else "omitted",
            "candidate_text_snippets": "included" if include_snippets else "omitted",
        },
        "content_parts": inspected_parts,
        "counts": counts,
        "media_files": media_files,
        "embedded_object_files": embedded_objects,
        "parse_errors": parse_errors,
        "limitations": [
            "Counts are structural signals, not a judgment of mathematical correctness.",
            "Text and LaTeX candidates are heuristic and may include code, URLs, citations, or prose.",
            "The scanner cannot identify whether an image contains a formula.",
            "The scanner cannot determine page count, visual clipping, fonts, or Word rendering fidelity.",
            "Source paths and candidate text snippets are omitted by default; include them only in local diagnostic reports.",
        ],
    }


def add_baseline_comparison(
    audit: dict[str, Any], baseline_path: Path, *, include_paths: bool = False
) -> None:
    try:
        baseline = json.loads(baseline_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Cannot read baseline JSON: {baseline_path}") from exc

    before_counts = baseline.get("counts")
    if not isinstance(before_counts, dict):
        raise ValueError("Baseline JSON does not contain a counts object")

    differences: dict[str, dict[str, int | None]] = {}
    for key in sorted(set(before_counts) | set(audit["counts"])):
        before = before_counts.get(key)
        after = audit["counts"].get(key)
        if before != after:
            differences[key] = {"before": before, "after": after}

    comparison = {
        "baseline_sha256": hashlib.sha256(baseline_path.read_bytes()).hexdigest(),
        "count_differences": differences,
    }
    if include_paths:
        comparison["baseline_path"] = str(baseline_path)
    audit["baseline_comparison"] = comparison


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a read-only OOXML/OMML inventory for a Word document."
    )
    parser.add_argument("document", type=Path, help="Input .docx or .docm file")
    parser.add_argument("--baseline", type=Path, help="Prior audit JSON for count comparison")
    parser.add_argument("--output", type=Path, help="Write JSON to this path instead of stdout")
    parser.add_argument("--compact", action="store_true", help="Write compact JSON")
    parser.add_argument(
        "--include-snippets",
        action="store_true",
        help="Include candidate text snippets; use only for local diagnostic reports",
    )
    parser.add_argument(
        "--include-paths",
        action="store_true",
        help="Include source and baseline paths; use only for local diagnostic reports",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        audit = make_audit(
            args.document,
            include_snippets=args.include_snippets,
            include_paths=args.include_paths,
        )
        if args.baseline:
            add_baseline_comparison(
                audit, args.baseline, include_paths=args.include_paths
            )
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    payload = json.dumps(
        audit,
        ensure_ascii=False,
        indent=None if args.compact else 2,
        sort_keys=False,
    )
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
