"""Batch and composite document building operations for Office MCP.

This module provides high-efficiency in-memory document building, allowing
full documents with typography, margins, line spacing, headers, footers,
tables, lists, and headings to be generated in a single pass.
"""

from __future__ import annotations

import re
from typing import Any
from docx.document import Document as DocumentType
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from word_mcp.core.styles import ensure_heading_style

ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

PRESETS = {
    "modern": {
        "font_name": "Calibri",
        "font_size": 11.0,
        "line_spacing": 1.15,
        "space_after": 6.0,
        "margins": {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5},
    },
    "formal": {
        "font_name": "Times New Roman",
        "font_size": 12.0,
        "line_spacing": 1.5,
        "space_after": 6.0,
        "margins": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5},
    },
    "academic": {
        "font_name": "Times New Roman",
        "font_size": 12.0,
        "line_spacing": 2.0,
        "space_after": 0.0,
        "margins": {"top": 2.54, "bottom": 2.54, "left": 2.54, "right": 2.54},
    },
    "compact": {
        "font_name": "Arial",
        "font_size": 10.0,
        "line_spacing": 1.0,
        "space_after": 4.0,
        "margins": {"top": 2.0, "bottom": 2.0, "left": 2.0, "right": 2.0},
    },
}


def apply_page_number_field(paragraph: Any) -> None:
    """Add a native Word PAGE field to a paragraph."""
    run = paragraph.add_run(" ")
    fld_simple = OxmlElement("w:fldSimple")
    fld_simple.set(qn("w:instr"), "PAGE")
    run._r.append(fld_simple)


def apply_document_settings(
    doc: DocumentType,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    orientation: str | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    preset: str | None = None,
) -> dict[str, Any]:
    """Configure document-wide styles, page setup, header, and footer."""
    config = dict(PRESETS.get(preset or "modern", PRESETS["modern"]))

    if font_name is not None:
        config["font_name"] = font_name
    if font_size is not None:
        config["font_size"] = font_size
    if line_spacing is not None:
        config["line_spacing"] = line_spacing
    if space_after is not None:
        config["space_after"] = space_after
    if margins is not None:
        config["margins"] = margins

    # Configure Normal style
    try:
        normal_style = doc.styles["Normal"]
        normal_font = normal_style.font
        normal_font.name = config["font_name"]
        normal_font.size = Pt(config["font_size"])

        p_format = normal_style.paragraph_format
        p_format.line_spacing = config["line_spacing"]
        p_format.space_after = Pt(config["space_after"])
    except (KeyError, AttributeError):
        pass

    # Ensure heading styles exist
    ensure_heading_style(doc)

    # Configure sections (margins, orientation, headers/footers)
    for section in doc.sections:
        m = config.get("margins", {})
        if "top" in m:
            section.top_margin = Cm(float(m["top"]))
        if "bottom" in m:
            section.bottom_margin = Cm(float(m["bottom"]))
        if "left" in m:
            section.left_margin = Cm(float(m["left"]))
        if "right" in m:
            section.right_margin = Cm(float(m["right"]))

        if orientation == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            new_width, new_height = section.page_height, section.page_width
            section.page_width = new_width
            section.page_height = new_height

        if header_text:
            header_para = section.header.paragraphs[0]
            header_para.text = header_text
            header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if len(header_para.runs) > 0:
                header_para.runs[0].font.size = Pt(8.5)
                header_para.runs[0].font.color.rgb = RGBColor(128, 128, 128)

        if footer_text or page_numbers:
            footer_para = section.footer.paragraphs[0]
            footer_para.text = footer_text or ""
            footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if page_numbers:
                if footer_text:
                    footer_para.add_run(" | Pág. ")
                apply_page_number_field(footer_para)
            if len(footer_para.runs) > 0:
                footer_para.runs[0].font.size = Pt(8.5)
                footer_para.runs[0].font.color.rgb = RGBColor(128, 128, 128)

    return config


def add_inline_formatted_text(paragraph: Any, text: str) -> None:
    """Parse inline **bold** and *italic* markdown in paragraph text."""
    # Pattern to match bold (**text**) or italic (*text*)
    tokens = re.split(r"(\*\*.*?\*\*|\*.*?\*)", text)
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) >= 4:
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        else:
            paragraph.add_run(token)


def render_element(doc: DocumentType, elem: dict[str, Any]) -> str:
    """Render a single content element dictionary into the docx Document.

    Returns:
        Type of rendered element.
    """
    elem_type = elem.get("type", "paragraph").lower()

    if elem_type == "heading":
        level = int(elem.get("level", 1))
        text = elem.get("text", "")
        p = doc.add_heading(text, level=min(max(level, 1), 9))
        align = elem.get("alignment")
        if align in ALIGNMENT_MAP:
            p.alignment = ALIGNMENT_MAP[align]
        return "heading"

    elif elem_type == "paragraph":
        text = elem.get("text", "")
        p = doc.add_paragraph()
        add_inline_formatted_text(p, text)

        align = elem.get("alignment")
        if align in ALIGNMENT_MAP:
            p.alignment = ALIGNMENT_MAP[align]

        if "space_after" in elem:
            p.paragraph_format.space_after = Pt(float(elem["space_after"]))
        if "line_spacing" in elem:
            p.paragraph_format.line_spacing = float(elem["line_spacing"])
        if elem.get("bold"):
            for run in p.runs:
                run.bold = True
        if elem.get("italic"):
            for run in p.runs:
                run.italic = True
        return "paragraph"

    elif elem_type == "table":
        headers = elem.get("headers", [])
        rows_data = elem.get("rows", [])
        style = elem.get("style", "Table Grid")

        total_rows = len(rows_data) + (1 if headers else 0)
        total_cols = len(headers) if headers else (len(rows_data[0]) if rows_data else 1)

        table = doc.add_table(rows=total_rows, cols=total_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        try:
            table.style = style
        except Exception:
            pass

        current_row = 0
        if headers:
            hdr_cells = table.rows[0].cells
            for idx, h_text in enumerate(headers[:total_cols]):
                hdr_cells[idx].text = str(h_text)
                for run in hdr_cells[idx].paragraphs[0].runs:
                    run.bold = True
            current_row = 1

        for r_data in rows_data:
            if current_row >= len(table.rows):
                break
            row_cells = table.rows[current_row].cells
            for c_idx, cell_value in enumerate(r_data[:total_cols]):
                row_cells[c_idx].text = str(cell_value)
            current_row += 1

        # Add empty spacing paragraph after table
        space_p = doc.add_paragraph()
        space_p.paragraph_format.space_after = Pt(4)
        return "table"

    elif elem_type in ("bullet_list", "bullet", "list"):
        items = elem.get("items", [])
        if isinstance(items, str):
            items = [items]
        for item in items:
            p = doc.add_paragraph(style="List Bullet")
            add_inline_formatted_text(p, str(item))
        return "bullet_list"

    elif elem_type in ("numbered_list", "number_list"):
        items = elem.get("items", [])
        if isinstance(items, str):
            items = [items]
        for item in items:
            p = doc.add_paragraph(style="List Number")
            add_inline_formatted_text(p, str(item))
        return "numbered_list"

    elif elem_type == "page_break":
        doc.add_page_break()
        return "page_break"

    elif elem_type in ("callout", "quote"):
        text = elem.get("text", "")
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.0)
        p.paragraph_format.right_indent = Cm(1.0)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        add_inline_formatted_text(p, text)
        for run in p.runs:
            run.italic = True
        return "callout"

    return "unknown"


def core_build_document(
    doc: DocumentType,
    title: str | None = None,
    author: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    orientation: str | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    preset: str | None = None,
    elements: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Build a complete Word document in memory in a single pass.

    Returns:
        Summary dictionary with counts of elements created.
    """
    # Core properties
    if title:
        doc.core_properties.title = title
    if author:
        doc.core_properties.author = author

    # Apply typography, margins, header, and footer
    config_applied = apply_document_settings(
        doc=doc,
        font_name=font_name,
        font_size=font_size,
        line_spacing=line_spacing,
        space_after=space_after,
        margins=margins,
        orientation=orientation,
        header_text=header_text,
        footer_text=footer_text,
        page_numbers=page_numbers,
        preset=preset,
    )

    counts: dict[str, int] = {}
    if elements:
        for elem in elements:
            if isinstance(elem, dict):
                t = render_element(doc, elem)
                counts[t] = counts.get(t, 0) + 1

    total_created = sum(counts.values())
    summary_parts = [f"{v} {k}s" for k, v in counts.items()]
    summary = ", ".join(summary_parts) if summary_parts else "Document layout configured"

    return {
        "status": "success",
        "total_elements": total_created,
        "element_counts": counts,
        "config_applied": config_applied,
        "summary": summary,
    }


def parse_markdown_to_elements(markdown_text: str) -> list[dict[str, Any]]:
    """Convert standard markdown text into a list of structured document elements."""
    elements: list[dict[str, Any]] = []
    lines = markdown_text.splitlines()
    in_table = False
    table_headers: list[str] = []
    table_rows: list[list[str]] = []

    def flush_table() -> None:
        nonlocal in_table, table_headers, table_rows
        if in_table and table_headers:
            elements.append({
                "type": "table",
                "headers": table_headers,
                "rows": table_rows,
                "style": "Table Grid",
            })
        in_table = False
        table_headers = []
        table_rows = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            flush_table()
            continue

        # Check for table line
        if line.startswith("|") and line.endswith("|"):
            cells = [c.strip() for c in line.split("|")[1:-1]]
            # Check if divider line like |---|---|
            if all(re.match(r"^:?-+:?$", c) for c in cells):
                continue
            if not in_table:
                in_table = True
                table_headers = cells
            else:
                table_rows.append(cells)
            continue
        else:
            flush_table()

        # Page break
        if line in ("---", "***", "___"):
            elements.append({"type": "page_break"})
            continue

        # Headings
        if line.startswith("#"):
            match = re.match(r"^(#{1,6})\s+(.*)$", line)
            if match:
                level = len(match.group(1))
                heading_text = match.group(2).strip()
                elements.append({"type": "heading", "level": level, "text": heading_text})
                continue

        # Bullet list
        if line.startswith(("- ", "* ", "+ ")):
            elements.append({"type": "bullet_list", "items": [line[2:].strip()]})
            continue

        # Numbered list
        num_match = re.match(r"^\d+\.\s+(.*)$", line)
        if num_match:
            elements.append({"type": "numbered_list", "items": [num_match.group(1).strip()]})
            continue

        # Blockquote / callout
        if line.startswith("> "):
            elements.append({"type": "callout", "text": line[2:].strip()})
            continue

        # Standard paragraph
        elements.append({"type": "paragraph", "text": line})

    flush_table()
    return elements


def core_create_from_markdown(
    doc: DocumentType,
    markdown_content: str,
    title: str | None = None,
    author: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    preset: str | None = None,
) -> dict[str, Any]:
    """Parse Markdown content and build Word document in a single pass."""
    elements = parse_markdown_to_elements(markdown_content)
    return core_build_document(
        doc=doc,
        title=title,
        author=author,
        font_name=font_name,
        font_size=font_size,
        line_spacing=line_spacing,
        space_after=space_after,
        margins=margins,
        header_text=header_text,
        footer_text=footer_text,
        page_numbers=page_numbers,
        preset=preset,
        elements=elements,
    )
