import pytest
from docx import Document
from docx.shared import Pt, RGBColor
from mcp_word.core.styles import StyleSettings, ensure_heading_style, DEFAULT_SETTINGS
from mcp_word.core.document import core_create_document
import os

def test_style_settings_defaults():
    settings = StyleSettings()
    assert settings.font_name == "Calibri"
    assert settings.get_heading_size(1) == Pt(16)
    assert settings.get_heading_size(2) == Pt(14)
    assert settings.get_heading_size(3) == Pt(12)

def test_style_settings_overrides():
    custom_headings = {1: Pt(24), 2: Pt(20)}
    custom_colors = {"brand": RGBColor(10, 20, 30)}
    settings = StyleSettings(
        font_name="Arial",
        font_size=Pt(12),
        heading_sizes=custom_headings,
        colors=custom_colors
    )
    assert settings.font_name == "Arial"
    assert settings.get_heading_size(1) == Pt(24)
    assert settings.get_heading_size(2) == Pt(20)
    assert settings.get_heading_size(3) == Pt(12)  # Default fallback
    assert settings.parse_color("brand") == RGBColor(10, 20, 30)

def test_ensure_heading_style():
    doc = Document()
    custom_size = Pt(33)
    custom_settings = StyleSettings(heading_sizes={1: custom_size})
    
    ensure_heading_style(doc, settings=custom_settings)
    
    style = doc.styles["Heading 1"]
    assert style.font.size == custom_size
    
    style2 = doc.styles["Heading 2"]
    assert style2.font.size == Pt(14) # Default from StyleSettings

def test_core_create_document_with_custom_styles(tmp_path):
    filename = str(tmp_path / "test_custom_styles.docx")
    custom_settings = {
        "font_name": "Times New Roman",
        "heading_sizes": {1: Pt(18), 2: Pt(16)}
    }
    
    core_create_document(filename, style_settings=custom_settings)
    
    assert os.path.exists(filename)
    doc = Document(filename)
    
    h1_style = doc.styles["Heading 1"]
    assert h1_style.font.name == "Times New Roman"
    assert h1_style.font.size == Pt(18)
    
    h2_style = doc.styles["Heading 2"]
    assert h2_style.font.size == Pt(16)

if __name__ == "__main__":
    # If running manually
    import sys
    pytest.main([__file__])
