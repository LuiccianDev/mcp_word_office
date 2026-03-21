"""
Style-related functionality for Word Document Server.

This module provides utilities for managing and applying styles in Word documents,
including creating custom styles, ensuring default styles exist, and applying
consistent formatting.
"""

# Standard library imports
from enum import Enum
from typing import Any, Dict, Optional, Union

# Third-party imports
from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Length, Pt, RGBColor


class StyleType(str, Enum):
    """Supported style types for document elements."""

    PARAGRAPH = "paragraph"
    CHARACTER = "character"
    TABLE = "table"
    NUMBERING = "numbering"


class StyleSettings:
    """Configuration for document styles and default values."""

    def __init__(
        self,
        font_name: str = "Calibri",
        font_size: Union[int, Length] = Pt(11),
        heading_sizes: Optional[Dict[int, Length]] = None,
        colors: Optional[Dict[str, RGBColor]] = None,
        default_color: RGBColor = RGBColor(0, 0, 0),
    ):
        self.font_name = font_name
        self.font_size = font_size
        self.default_color = default_color

        # Default heading sizes mapping level -> size
        self.heading_sizes = {
            1: Pt(16),
            2: Pt(14),
            3: Pt(12),
            4: Pt(12),
            5: Pt(12),
            6: Pt(12),
            7: Pt(12),
            8: Pt(12),
            9: Pt(12),
        }
        if heading_sizes:
            self.heading_sizes.update(heading_sizes)

        # Extensible color map
        self.colors = {
            "red": RGBColor(255, 0, 0),
            "blue": RGBColor(0, 0, 255),
            "green": RGBColor(0, 128, 0),
            "yellow": RGBColor(255, 255, 0),
            "black": RGBColor(0, 0, 0),
            "gray": RGBColor(128, 128, 128),
            "white": RGBColor(255, 255, 255),
            "purple": RGBColor(128, 0, 128),
            "orange": RGBColor(255, 165, 0),
        }
        if colors:
            self.colors.update(colors)

    def get_heading_size(self, level: int) -> Length:
        """Get the font size for a specific heading level."""
        return self.heading_sizes.get(level, self.heading_sizes.get(3, Pt(12)))

    def parse_color(self, value: Any) -> RGBColor:
        """Parse color value to RGBColor using the defined color map."""
        if value is None:
            return self.default_color
        if isinstance(value, RGBColor):
            return value
        if isinstance(value, str):
            color_lower = value.lower()
            if color_lower in self.colors:
                return self.colors[color_lower]
            try:
                # Attempt to parse from hex string or similar
                return RGBColor.from_string(value)
            except (ValueError, AttributeError):
                pass
        return self.default_color


# Global default settings
DEFAULT_SETTINGS = StyleSettings()


class FontProperties:
    """Model for font-related style properties."""

    def __init__(
        self,
        name: Optional[str] = None,
        size: Optional[Union[int, Length]] = None,
        bold: bool = False,
        italic: bool = False,
        color: Optional[Union[str, RGBColor]] = None,
        settings: StyleSettings = DEFAULT_SETTINGS,
    ):
        self.name = name or settings.font_name
        self.size = size if size is not None else settings.font_size
        self.bold = bold
        self.italic = italic
        self.color = settings.parse_color(color)


class ParagraphProperties:
    """Model for paragraph-related style properties."""

    def __init__(
        self,
        alignment: Optional[int] = None,
        spacing: Optional[float] = None,
        space_before: Optional[Length] = None,
        space_after: Optional[Length] = None,
    ):
        self.alignment = alignment
        self.spacing = spacing
        self.space_before = space_before
        self.space_after = space_after


def ensure_heading_style(doc: Document, settings: StyleSettings = DEFAULT_SETTINGS) -> None:
    """
    Ensure default heading styles exist in the document and match settings.

    Args:
        doc: The document to ensure styles for.
        settings: Style settings to use for formatting.
    """
    for level in range(1, 10):
        style_name = f"Heading {level}"
        try:
            if style_name not in doc.styles:
                style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            else:
                style = doc.styles[style_name]

            font = style.font
            font.name = settings.font_name
            font.bold = True
            font.size = settings.get_heading_size(level)
        except Exception as e:
            # Log error but continue with other styles
            print(f"Warning: Failed to ensure {style_name} style: {e}")


def ensure_table_style(doc: Document) -> None:
    """
    Ensure default table style exists in the document.

    Args:
        doc: The document to ensure the table style for.
    """
    style_name = "Table Grid"
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
    paragraph_props: Optional[Dict[str, Any]] = None,
    settings: StyleSettings = DEFAULT_SETTINGS,
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
        settings: Style settings to use for defaults.

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
            # Pass settings to FontProperties
            font_props_obj = FontProperties(**font_props, settings=settings)

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
        if paragraph_props and hasattr(style, "paragraph_format"):
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
