#!/usr/bin/env python
"""Create the bundled, sanitized party-building DOCX standard template."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


STYLE_PROFILES = {
    "党建标题": {
        "font": "方正小标宋简体",
        "size_pt": 22,
        "bold": False,
        "alignment": "CENTER",
        "first_line_indent_pt": None,
        "line_spacing_pt": 36,
        "keep_with_next": True,
    },
    "党建标题1": {
        "font": "宋体",
        "size_pt": 22,
        "bold": True,
        "color": "333333",
        "alignment": "CENTER",
        "first_line_indent_pt": None,
        "line_spacing_pt": 36,
        "keep_with_next": True,
    },
    "党建正文": {
        "font": "仿宋_GB2312",
        "size_pt": 16,
        "bold": False,
        "alignment": "JUSTIFY",
        "first_line_indent_pt": 32,
        "line_spacing_pt": 31,
    },
    "全文一级大标题": {
        "font": "黑体",
        "size_pt": 22,
        "bold": True,
        "alignment": "CENTER",
        "first_line_indent_pt": None,
        "line_spacing_pt": 32,
        "keep_with_next": True,
    },
    "正文一级标题": {
        "font": "黑体",
        "size_pt": 18,
        "bold": True,
        "alignment": "JUSTIFY",
        "first_line_indent_pt": 36,
        "line_spacing_pt": 28,
        "keep_with_next": True,
    },
    "正文二级标题": {
        "font": "黑体",
        "size_pt": 17,
        "bold": True,
        "alignment": "JUSTIFY",
        "first_line_indent_pt": 34,
        "line_spacing_pt": 28,
        "keep_with_next": True,
    },
    "正文 文本": {
        "font": "宋体",
        "size_pt": 16,
        "bold": False,
        "alignment": "JUSTIFY",
        "first_line_indent_pt": 32,
        "line_spacing_pt": 28,
    },
}


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def import_dependencies():
    try:
        from docx import Document
        from docx.enum.style import WD_STYLE_TYPE
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Cm, Pt, RGBColor
    except ImportError as exc:
        raise SystemExit(
            "Missing dependency. Run this script with the bundled workspace "
            "Python that includes python-docx."
        ) from exc
    return {
        "Document": Document,
        "WD_STYLE_TYPE": WD_STYLE_TYPE,
        "WD_ALIGN_PARAGRAPH": WD_ALIGN_PARAGRAPH,
        "WD_LINE_SPACING": WD_LINE_SPACING,
        "OxmlElement": OxmlElement,
        "qn": qn,
        "Cm": Cm,
        "Pt": Pt,
        "RGBColor": RGBColor,
    }


def set_style_font(style, font_name: str, size_pt: float, bold: bool, deps) -> None:
    style.font.name = font_name
    style.font.size = deps["Pt"](size_pt)
    style.font.bold = bold
    run_properties = style.element.get_or_add_rPr()
    fonts = run_properties.get_or_add_rFonts()
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(deps["qn"](f"w:{attribute}"), font_name)


def configure_standard_template(document: object, deps) -> None:
    section = document.sections[0]
    section.page_width = deps["Cm"](21.0)
    section.page_height = deps["Cm"](29.7)
    section.top_margin = deps["Cm"](2.54)
    section.bottom_margin = deps["Cm"](2.54)
    section.left_margin = deps["Cm"](3.17)
    section.right_margin = deps["Cm"](3.17)
    section.header_distance = deps["Cm"](1.5)
    section.footer_distance = deps["Cm"](1.75)

    section_properties = section._sectPr
    document_grid = section_properties.find(deps["qn"]("w:docGrid"))
    if document_grid is None:
        document_grid = deps["OxmlElement"]("w:docGrid")
        section_properties.append(document_grid)
    document_grid.set(deps["qn"]("w:type"), "lines")
    document_grid.set(deps["qn"]("w:linePitch"), "312")

    normal = document.styles["Normal"]
    set_style_font(normal, "宋体", 10.5, False, deps)
    normal.paragraph_format.space_before = deps["Pt"](0)
    normal.paragraph_format.space_after = deps["Pt"](0)
    normal.paragraph_format.alignment = deps["WD_ALIGN_PARAGRAPH"].JUSTIFY

    for style_name, profile in STYLE_PROFILES.items():
        if style_name in document.styles:
            style = document.styles[style_name]
        else:
            style = document.styles.add_style(
                style_name, deps["WD_STYLE_TYPE"].PARAGRAPH
            )
        style.base_style = normal
        set_style_font(
            style,
            profile["font"],
            profile["size_pt"],
            profile["bold"],
            deps,
        )
        if profile.get("color"):
            style.font.color.rgb = deps["RGBColor"].from_string(profile["color"])

        paragraph_format = style.paragraph_format
        paragraph_format.alignment = getattr(
            deps["WD_ALIGN_PARAGRAPH"], profile["alignment"]
        )
        paragraph_format.space_before = deps["Pt"](0)
        paragraph_format.space_after = deps["Pt"](0)
        paragraph_format.line_spacing_rule = deps["WD_LINE_SPACING"].EXACTLY
        paragraph_format.line_spacing = deps["Pt"](profile["line_spacing_pt"])
        if profile["first_line_indent_pt"] is None:
            paragraph_format.first_line_indent = deps["Pt"](0)
        else:
            paragraph_format.first_line_indent = deps["Pt"](
                profile["first_line_indent_pt"]
            )
        if profile.get("keep_with_next"):
            paragraph_format.keep_with_next = True

    properties = document.core_properties
    properties.title = "党建材料标准模板"
    properties.subject = "通用、可复用的党建记录DOCX模板"
    properties.author = ""
    properties.last_modified_by = ""
    properties.comments = ""


def build_template(output_path: Path) -> Path:
    deps = import_dependencies()
    document = deps["Document"]()
    configure_standard_template(document, deps)
    for paragraph in list(document.paragraphs):
        paragraph._element.getparent().remove(paragraph._element)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)

    reopened = deps["Document"](output_path)
    missing = sorted(
        style_name
        for style_name in STYLE_PROFILES
        if style_name not in reopened.styles
    )
    if missing:
        output_path.unlink(missing_ok=True)
        raise SystemExit(f"Standard template is missing styles: {missing}")
    return output_path


def main() -> None:
    configure_stdout()
    default_output = (
        Path(__file__).resolve().parents[1]
        / "templates"
        / "standard-party-template.docx"
    )
    parser = argparse.ArgumentParser(
        description="Create the sanitized standard party-building DOCX template."
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=default_output,
        help="Template output path.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace an existing template.",
    )
    args = parser.parse_args()

    output = args.output.resolve()
    if output.exists() and not args.overwrite:
        raise SystemExit(f"Template already exists; pass --overwrite: {output}")
    path = build_template(output)
    print(f"Created standard template: {path}")


if __name__ == "__main__":
    main()
