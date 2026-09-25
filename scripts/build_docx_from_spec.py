#!/usr/bin/env python
"""Build DOCX records by copying local templates and replacing body paragraphs."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path


ALIGNMENTS = {
    "left": "LEFT",
    "center": "CENTER",
    "right": "RIGHT",
    "justify": "JUSTIFY",
}


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def import_dependencies():
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt
    except ImportError as exc:
        raise SystemExit(
            "Missing dependency. Run this script with the bundled workspace "
            "Python that includes python-docx."
        ) from exc
    return Document, WD_ALIGN_PARAGRAPH, Pt


def load_spec(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"Cannot read JSON spec: {exc}") from exc
    if not isinstance(data, dict) or not isinstance(data.get("documents"), list):
        raise SystemExit("Spec must be an object with a non-empty 'documents' list.")
    if not data["documents"]:
        raise SystemExit("Spec contains no documents.")
    return data


def resolve_path(value: str, base_dir: Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else (base_dir / path).resolve()


def validate_document_spec(spec: dict, base_dir: Path, index: int) -> dict:
    prefix = f"documents[{index}]"
    if not isinstance(spec, dict):
        raise SystemExit(f"{prefix} must be an object.")
    for field in ("template", "filename", "paragraphs"):
        if field not in spec:
            raise SystemExit(f"{prefix} is missing '{field}'.")
    if not isinstance(spec["paragraphs"], list) or not spec["paragraphs"]:
        raise SystemExit(f"{prefix}.paragraphs must be a non-empty list.")

    template = resolve_path(str(spec["template"]), base_dir)
    if not template.is_file() or template.suffix.lower() != ".docx":
        raise SystemExit(f"{prefix}.template is not a DOCX file: {template}")

    filename = str(spec["filename"])
    if Path(filename).name != filename or not filename.lower().endswith(".docx"):
        raise SystemExit(f"{prefix}.filename must be a DOCX basename.")

    paragraphs = []
    for paragraph_index, paragraph in enumerate(spec["paragraphs"]):
        paragraph_prefix = f"{prefix}.paragraphs[{paragraph_index}]"
        if not isinstance(paragraph, dict):
            raise SystemExit(f"{paragraph_prefix} must be an object.")
        if not str(paragraph.get("style", "")).strip():
            raise SystemExit(f"{paragraph_prefix} is missing 'style'.")
        if not str(paragraph.get("text", "")).strip():
            raise SystemExit(f"{paragraph_prefix} is missing non-empty 'text'.")
        alignment = paragraph.get("alignment")
        if alignment is not None and alignment not in ALIGNMENTS:
            raise SystemExit(
                f"{paragraph_prefix}.alignment must be one of "
                f"{sorted(ALIGNMENTS)}."
            )
        paragraphs.append(
            {
                "style": str(paragraph["style"]),
                "text": str(paragraph["text"]),
                "bold": bool(paragraph.get("bold", False)),
                "alignment": alignment,
                "first_line_indent_pt": paragraph.get("first_line_indent_pt"),
            }
        )

    return {
        "template": template,
        "filename": filename,
        "title": str(spec.get("title", Path(filename).stem)),
        "paragraphs": paragraphs,
    }


def clear_paragraphs(document) -> None:
    for paragraph in list(document.paragraphs):
        paragraph._element.getparent().remove(paragraph._element)


def build_document(
    Document,
    WD_ALIGN_PARAGRAPH,
    Pt,
    document_spec: dict,
    output_dir: Path,
    overwrite: bool,
    dry_run: bool,
) -> Path:
    output_path = output_dir / document_spec["filename"]
    if output_path.exists() and not overwrite:
        raise SystemExit(f"Output already exists; pass --overwrite: {output_path}")
    if dry_run:
        return output_path

    output_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(document_spec["template"], output_path)
    document = Document(output_path)

    missing_styles = sorted(
        {
            paragraph["style"]
            for paragraph in document_spec["paragraphs"]
            if paragraph["style"] not in document.styles
        }
    )
    if missing_styles:
        output_path.unlink(missing_ok=True)
        raise SystemExit(
            f"Template {document_spec['template']} is missing styles: "
            f"{missing_styles}"
        )

    clear_paragraphs(document)
    for paragraph_spec in document_spec["paragraphs"]:
        paragraph = document.add_paragraph(style=paragraph_spec["style"])
        run = paragraph.add_run(paragraph_spec["text"])
        run.bold = paragraph_spec["bold"]

        alignment = paragraph_spec["alignment"]
        if alignment:
            paragraph.paragraph_format.alignment = getattr(
                WD_ALIGN_PARAGRAPH, ALIGNMENTS[alignment]
            )

        indent_pt = paragraph_spec["first_line_indent_pt"]
        if indent_pt is not None:
            paragraph.paragraph_format.first_line_indent = Pt(float(indent_pt))
        elif paragraph_spec["bold"]:
            paragraph.paragraph_format.first_line_indent = Pt(0)

    document.core_properties.title = document_spec["title"]
    document.save(output_path)
    return output_path


def main() -> None:
    configure_stdout()
    parser = argparse.ArgumentParser(
        description="Build DOCX records from a template-inheritance JSON spec."
    )
    parser.add_argument("--spec", required=True, type=Path)
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Override output_dir from the spec.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing output files.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate the spec without writing DOCX files.",
    )
    args = parser.parse_args()

    spec_path = args.spec.resolve()
    raw_spec = load_spec(spec_path)
    base_dir = spec_path.parent
    output_value = str(args.output_dir or raw_spec.get("output_dir", "output"))
    output_dir = resolve_path(output_value, base_dir)

    document_specs = [
        validate_document_spec(document, base_dir, index)
        for index, document in enumerate(raw_spec["documents"])
    ]

    Document, WD_ALIGN_PARAGRAPH, Pt = import_dependencies()
    outputs = [
        build_document(
            Document,
            WD_ALIGN_PARAGRAPH,
            Pt,
            document_spec,
            output_dir,
            args.overwrite,
            args.dry_run,
        )
        for document_spec in document_specs
    ]

    action = "Validated" if args.dry_run else "Created"
    print(f"{action} {len(outputs)} document(s):")
    for output in outputs:
        print(output)


if __name__ == "__main__":
    main()
