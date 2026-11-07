"""Tests for batch and composite document building tools."""

import asyncio
import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document

from word_mcp.tools import batch_tools, header_footer_tools, link_tools


class TestBatchTools(unittest.IsolatedAsyncioTestCase):
    """Test suite for high-efficiency batch tools."""

    def setUp(self) -> None:
        self.temp_dir = TemporaryDirectory()
        self.doc_path = str(Path(self.temp_dir.name) / "test_doc.docx")
        os.environ["MCP_ALLOWED_DIRECTORIES"] = self.temp_dir.name

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    async def test_build_document_complete(self) -> None:
        """Test building a complete document with typography, margins, header, footer, and elements in 1 pass."""
        elements = [
            {"type": "heading", "text": "1. Resumen Ejecutivo", "level": 1, "alignment": "center"},
            {
                "type": "paragraph",
                "text": "Este es el primer párrafo con texto en **negrita** y *cursiva*.",
                "alignment": "justify",
                "space_after": 8,
            },
            {
                "type": "table",
                "headers": ["Categoría", "Q1", "Q2"],
                "rows": [
                    ["Ingresos", "$100k", "$120k"],
                    ["Gastos", "$60k", "$70k"],
                ],
                "style": "Table Grid",
            },
            {"type": "page_break"},
            {"type": "heading", "text": "2. Plan de Acción", "level": 2},
            {"type": "bullet_list", "items": ["Acción 1", "Acción 2"]},
            {"type": "numbered_list", "items": ["Paso A", "Paso B"]},
            {"type": "callout", "text": "Nota importante de advertencia."},
        ]

        result = await batch_tools.build_document(
            filename=self.doc_path,
            title="Informe de Prueba",
            author="Tester",
            preset="formal",
            font_name="Times New Roman",
            font_size=12.0,
            line_spacing=1.5,
            space_after=6.0,
            margins={"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5},
            header_text="Header de Prueba",
            footer_text="Footer Confidencial",
            page_numbers=True,
            elements=elements,
        )

        self.assertEqual(result["status"], "success")
        self.assertTrue(os.path.exists(self.doc_path))
        self.assertGreater(result["total_elements"], 5)
        self.assertIn("headings", result["summary"])

        # Inspect resulting DOCX
        doc = Document(self.doc_path)
        self.assertEqual(doc.core_properties.title, "Informe de Prueba")
        self.assertEqual(doc.core_properties.author, "Tester")

        # Verify sections, margins, and header
        section = doc.sections[0]
        self.assertIn("Header de Prueba", section.header.paragraphs[0].text)
        self.assertIn("Footer Confidencial", section.footer.paragraphs[0].text)

        # Verify tables and paragraphs
        self.assertEqual(len(doc.tables), 1)
        self.assertEqual(len(doc.tables[0].rows), 3)  # header + 2 data rows
        self.assertEqual(doc.tables[0].rows[0].cells[0].text, "Categoría")

    async def test_create_from_markdown(self) -> None:
        """Test generating a styled document directly from Markdown text."""
        md_text = """# Título Principal
Este es un párrafo de introducción con **texto destacado**.

## Subsección 1
- Primer elemento de lista
- Segundo elemento de lista

| Producto | Cantidad | Precio |
|---|---|---|
| Widget A | 10 | $5.00 |
| Widget B | 25 | $12.50 |

> Nota destacada en bloque.

---
# Segunda Página
Texto después del salto de página.
"""
        result = await batch_tools.create_from_markdown(
            filename=self.doc_path,
            markdown_content=md_text,
            title="Documento Markdown",
            header_text="Doc Markdown",
            page_numbers=True,
            preset="modern",
        )

        self.assertEqual(result["status"], "success")
        self.assertTrue(os.path.exists(self.doc_path))

        doc = Document(self.doc_path)
        self.assertGreaterEqual(len(doc.paragraphs), 4)
        self.assertEqual(len(doc.tables), 1)
        self.assertEqual(doc.tables[0].rows[0].cells[0].text, "Producto")
        self.assertEqual(doc.tables[0].rows[1].cells[1].text, "10")

    async def test_batch_add_content(self) -> None:
        """Test appending content elements to an existing document."""
        # Create base doc
        doc = Document()
        doc.add_paragraph("Texto inicial.")
        doc.save(self.doc_path)

        elements = [
            {"type": "heading", "text": "Sección Anexa", "level": 2},
            {"type": "paragraph", "text": "Texto anexo en lote."},
        ]

        result = await batch_tools.batch_add_content(
            filename=self.doc_path,
            elements=elements,
        )

        self.assertEqual(result["status"], "success")
        self.assertEqual(result["total_elements"], 2)

        doc = Document(self.doc_path)
        texts = [p.text for p in doc.paragraphs]
        self.assertIn("Sección Anexa", texts)
        self.assertIn("Texto anexo en lote.", texts)

    async def test_parameter_alias_tolerance(self) -> None:
        """Test that tools accept 'filename' or 'file_path' transparently without errors."""
        doc = Document()
        doc.add_paragraph("Párrafo de prueba para bookmark.")
        doc.save(self.doc_path)

        # Call add_header using filename instead of file_path
        res_header = await header_footer_tools.add_header(
            filename=self.doc_path,
            text="Header con alias filename",
        )
        self.assertEqual(res_header["status"], "success")

        # Call add_footer using filename instead of file_path
        res_footer = await header_footer_tools.add_footer(
            filename=self.doc_path,
            text="Footer con alias filename",
        )
        self.assertEqual(res_footer["status"], "success")

        # Call add_bookmark using filename instead of file_path
        res_bm = await link_tools.add_bookmark(
            filename=self.doc_path,
            paragraph_index=0,
            name="bm_test",
        )
        self.assertEqual(res_bm["status"], "success")


if __name__ == "__main__":
    unittest.main()
