"""
Style-related functionality for Word Document Server.

This module provides utilities for managing and applying styles in Word documents,
including creating custom styles, ensuring default styles exist, and applying
consistent formatting.
"""
# Standard library imports
from dataclasses import dataclass
from typing import Dict, Optional, Type, TypeVar, Any, Union, List
from enum import Enum

# Third-party imports
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor, Pt, Length
# Eliminada la dependencia de Pydantic

# Type variable for style type
T = TypeVar('T', bound='StyleBase')

# Constants for default values
DEFAULT_HEADING_1_SIZE = Pt(16)
DEFAULT_HEADING_2_SIZE = Pt(14)
DEFAULT_HEADING_SIZE = Pt(12)
DEFAULT_FONT_NAME = 'Calibri'
DEFAULT_FONT_SIZE = Pt(11)
DEFAULT_COLOR = RGBColor(0, 0, 0)  # Black

class StyleType(str, Enum):
    """Supported style types for document elements."""
    PARAGRAPH = 'paragraph'
    CHARACTER = 'character'
    TABLE = 'table'
    NUMBERING = 'numbering'

def _parse_color(value: Any) -> RGBColor:
    """Parse color value to RGBColor."""
    if value is None:
        return DEFAULT_COLOR
    if isinstance(value, RGBColor):
        return value
    if isinstance(value, str):
        color_map = {
            'red': RGBColor(255, 0, 0),
            'blue': RGBColor(0, 0, 255),
            'green': RGBColor(0, 128, 0),
            'yellow': RGBColor(255, 255, 0),
            'black': RGBColor(0, 0, 0),
            'gray': RGBColor(128, 128, 128),
            'white': RGBColor(255, 255, 255),
            'purple': RGBColor(128, 0, 128),
            'orange': RGBColor(255, 165, 0)
        }
        color_lower = value.lower()
        if color_lower in color_map:
            return color_map[color_lower]
        try:
            return RGBColor.from_string(value)
        except (ValueError, AttributeError):
            pass
    return DEFAULT_COLOR

class FontProperties:
    """Model for font-related style properties."""
    def __init__(
        self,
        name: str = DEFAULT_FONT_NAME,
        size: Optional[Union[int, Length]] = DEFAULT_FONT_SIZE,
        bold: bool = False,
        italic: bool = False,
        color: Optional[Union[str, RGBColor]] = None
    ):
        self.name = name
        self.size = size
        self.bold = bold
        self.italic = italic
        self.color = _parse_color(color)

class ParagraphProperties:
    """Model for paragraph-related style properties."""
    def __init__(
        self,
        alignment: Optional[int] = None,
        spacing: Optional[float] = None,
        space_before: Optional[Length] = None,
        space_after: Optional[Length] = None
    ):
        self.alignment = alignment
        self.spacing = spacing
        self.space_before = space_before
        self.space_after = space_after

def ensure_heading_style(doc: Document) -> None:
    """
    Ensure default heading styles exist in the document.

    Args:
        doc: The document to ensure styles for.
    """
    for level in range(1, 10):
        style_name = f'Heading {level}'
        if style_name not in doc.styles:
            try:
                style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                font = style.font
                font.name = 'Calibri'
                font.bold = True
                
                # Set appropriate font size based on heading level
                if level == 1:
                    font.size = DEFAULT_HEADING_1_SIZE
                elif level == 2:
                    font.size = DEFAULT_HEADING_2_SIZE
                else:
                    font.size = DEFAULT_HEADING_SIZE
            except Exception as e:
                # Log error but continue with other styles
                print(f"Warning: Failed to create {style_name} style: {e}")

def ensure_table_style(doc: Document) -> None:
    """
    Ensure default table style exists in the document.

    Args:
        doc: The document to ensure the table style for.
    """
    style_name = 'Table Grid'
    if style_name not in doc.styles:
        try:
            doc.styles.add_style(style_name, WD_STYLE_TYPE.TABLE)
        except Exception as e:
            print(f"Warning: Failed to create {style_name} style: {e}")

def create_style(
    doc: Document,
    style_name: str,
    style_type: WD_STYLE_TYPE = WD_STYLE_TYPE.PARAGRAPH,
    base_style: Optional[str] = None,
    font_props: Optional[Dict[str, Any]] = None,
    paragraph_props: Optional[Dict[str, Any]] = None
) -> Any:
    """
    Create or update a style in the document.

    Args:
        doc: The document to add the style to.
        style_name: Name of the style to create or update.
        style_type: Type of style to create.
        base_style: Optional name of the base style to inherit from.
        font_props: Dictionary of font properties.
        paragraph_props: Dictionary of paragraph properties.

    Returns:
        The created or existing style.
    """
    try:
        # Try to get existing style
        return doc.styles.get_by_id(style_name, style_type)
    except KeyError:
        # Create new style
        style = doc.styles.add_style(style_name, style_type)
        
        # Set base style if specified
        if base_style and base_style in doc.styles:
            style.base_style = doc.styles[base_style]
        
        # Apply font properties if provided
        if font_props:
            font = style.font
            font_props_obj = FontProperties(**font_props)
            
            if font_props_obj.name:
                font.name = font_props_obj.name
            if font_props_obj.size is not None:
                font.size = font_props_obj.size
            if font_props_obj.bold is not None:
                font.bold = font_props_obj.bold
            if font_props_obj.italic is not None:
                font.italic = font_props_obj.italic
            if font_props_obj.color is not None:
                font.color.rgb = font_props_obj.color
        
        # Apply paragraph properties if provided and style supports it
        if paragraph_props and hasattr(style, 'paragraph_format'):
            para_props = ParagraphProperties(**paragraph_props)
            para_format = style.paragraph_format
            
            if para_props.alignment is not None:
                para_format.alignment = para_props.alignment
            if para_props.spacing is not None:
                para_format.line_spacing = para_props.spacing
            if para_props.space_before is not None:
                para_format.space_before = para_props.space_before
            if para_props.space_after is not None:
                para_format.space_after = para_props.space_after
        
        return style
