#!/usr/bin/env python3
"""Rebuild a DOCX reference list from finalized, styled JSON entries."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Rebuild a Word reference list from a JSON array. Each array item may "
            "be a plain string or an object with a segments array containing "
            "{'text': str, 'italic': bool, 'bold': bool} objects."
        )
    )
    parser.add_argument("--input", required=True, type=Path, help="Source DOCX")
    parser.add_argument(
        "--references", required=True, type=Path, help="UTF-8 JSON reference array"
    )
    parser.add_argument("--output", required=True, type=Path, help="Output DOCX")
    parser.add_argument("--title", default="References", help="Centered heading")
    parser.add_argument(
        "--spacing",
        choices=("single", "double"),
        default="single",
        help="Reference paragraph line spacing",
    )
    parser.add_argument(
        "--hanging-indent",
        type=float,
        default=0.3,
        help="Hanging indent in inches (default: 0.3)",
    )
    return parser.parse_args()


def set_run_font(run: Any, *, size: float = 12.0) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attr}"), "Times New Roman")


def normalize_entry(item: Any, index: int) -> list[dict[str, Any]]:
    if isinstance(item, str):
        text = " ".join(item.split())
        if not text:
            raise ValueError(f"Reference {index} is empty")
        return [{"text": text, "italic": False, "bold": False}]

    if not isinstance(item, dict) or not isinstance(item.get("segments"), list):
        raise ValueError(
            f"Reference {index} must be a string or an object with a segments array"
        )

    normalized: list[dict[str, Any]] = []
    for segment_index, segment in enumerate(item["segments"], start=1):
        if not isinstance(segment, dict) or not isinstance(segment.get("text"), str):
            raise ValueError(
                f"Reference {index}, segment {segment_index} must contain string text"
            )
        raw_text = segment["text"]
        core_text = " ".join(raw_text.split())
        if not core_text:
            continue
        text = (
            (" " if raw_text[:1].isspace() else "")
            + core_text
            + (" " if raw_text[-1:].isspace() else "")
        )
        normalized.append(
            {
                "text": text,
                "italic": bool(segment.get("italic", False)),
                "bold": bool(segment.get("bold", False)),
            }
        )

    if not normalized:
        raise ValueError(f"Reference {index} has no nonempty segments")
    return normalized


def load_references(path: Path) -> list[list[dict[str, Any]]]:
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, list) or not data:
        raise ValueError("The reference JSON must be a nonempty array")
    return [normalize_entry(item, i) for i, item in enumerate(data, start=1)]


def clear_document_body(document: Any) -> None:
    body = document._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def clear_story(container: Any) -> None:
    for table in list(container.tables):
        table._element.getparent().remove(table._element)
    paragraphs = list(container.paragraphs)
    for paragraph in paragraphs[1:]:
        paragraph._element.getparent().remove(paragraph._element)
    if paragraphs:
        first = paragraphs[0]
        for child in list(first._element):
            if child.tag != qn("w:pPr"):
                first._element.remove(child)


def configure_document(document: Any) -> None:
    styles = document.styles
    normal = styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal._element.get_or_add_rPr().get_or_add_rFonts().set(
        qn("w:eastAsia"), "Times New Roman"
    )
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    if "Reference List" not in styles:
        ref_style = styles.add_style("Reference List", WD_STYLE_TYPE.PARAGRAPH)
    else:
        ref_style = styles["Reference List"]
    ref_style.base_style = normal

    for section in document.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.header_distance = Inches(0.5)
        section.footer_distance = Inches(0.5)
        clear_story(section.header)
        clear_story(section.footer)


def add_title(document: Any, title: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(12)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    paragraph.paragraph_format.keep_with_next = True
    run = paragraph.add_run(title)
    run.bold = True
    set_run_font(run)


def add_reference(
    document: Any,
    segments: list[dict[str, Any]],
    *,
    hanging_indent: float,
    spacing: str,
) -> None:
    paragraph = document.add_paragraph(style="Reference List")
    fmt = paragraph.paragraph_format
    fmt.left_indent = Inches(hanging_indent)
    fmt.first_line_indent = Inches(-hanging_indent)
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.keep_with_next = False
    fmt.keep_together = False
    fmt.widow_control = True
    fmt.line_spacing_rule = (
        WD_LINE_SPACING.DOUBLE if spacing == "double" else WD_LINE_SPACING.SINGLE
    )

    for segment in segments:
        run = paragraph.add_run(segment["text"])
        run.italic = segment["italic"]
        run.bold = segment["bold"]
        set_run_font(run)


def main() -> int:
    args = parse_args()
    if args.hanging_indent <= 0:
        raise ValueError("--hanging-indent must be greater than zero")
    if not args.input.is_file():
        raise FileNotFoundError(f"Input DOCX not found: {args.input}")
    if args.input.suffix.lower() != ".docx" or args.output.suffix.lower() != ".docx":
        raise ValueError("Both --input and --output must be .docx files")

    references = load_references(args.references)
    document = Document(str(args.input))
    clear_document_body(document)
    configure_document(document)
    add_title(document, args.title)
    for segments in references:
        add_reference(
            document,
            segments,
            hanging_indent=args.hanging_indent,
            spacing=args.spacing,
        )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    document.core_properties.title = args.title
    document.save(str(args.output))
    print(f"Saved {len(references)} references to {args.output}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, json.JSONDecodeError) as error:
        print(f"error: {error}", file=sys.stderr)
        raise SystemExit(2)
