# Presets and Styling Reference for `word-mcp`

## Available Presets

When calling `build_document` or formatting tools, choose the preset that matches the document objective:

### 1. `modern` (Default)
- **Body Font**: Calibri, 11 pt, Color `#262626`
- **Headings**: Calibri Light (Heading 1: 18 pt bold `#0F4C81`, Heading 2: 14 pt bold `#2C5E8A`, Heading 3: 12 pt semi-bold)
- **Line Spacing**: 1.15 lines
- **Paragraph Spacing**: 6 pt after
- **Margins**: Top: 2.5 cm, Bottom: 2.5 cm, Left: 2.5 cm, Right: 2.5 cm
- **Use Case**: Business proposals, executive reports, status updates, corporate communications.

### 2. `formal`
- **Body Font**: Times New Roman, 12 pt, Color `#000000`
- **Headings**: Times New Roman (Heading 1: 16 pt bold `#1A1A1A`, Heading 2: 13 pt bold, Heading 3: 12 pt bold)
- **Line Spacing**: 1.50 lines
- **Paragraph Spacing**: 6 pt after
- **Margins**: Top: 2.5 cm, Bottom: 2.5 cm, Left: 3.0 cm, Right: 2.5 cm (standard binding margin)
- **Use Case**: Legal contracts, corporate acts, formal requests, institutional memoranda.

### 3. `academic`
- **Body Font**: Times New Roman, 12 pt, Color `#000000`
- **Headings**: Times New Roman, bold
- **Line Spacing**: 2.00 lines (Double spaced per APA / MLA standards)
- **Paragraph Spacing**: 0 pt after (indented first line)
- **Margins**: 2.54 cm (1.0 inch) on all 4 sides
- **Use Case**: University papers, scientific articles, literature reviews, theses.

### 4. `compact`
- **Body Font**: Arial, 10 pt
- **Headings**: Arial, bold (Heading 1: 14 pt, Heading 2: 12 pt, Heading 3: 10 pt)
- **Line Spacing**: 1.00 lines (Single spaced)
- **Paragraph Spacing**: 4 pt after
- **Margins**: 2.0 cm on all sides
- **Use Case**: Invoices, product spec sheets, brief internal notes, cheat sheets.

## Custom Margins

You can override preset margins in `build_document` using the `margins` object:
```json
"margins": {
  "top": 2.5,
  "bottom": 2.5,
  "left": 3.0,
  "right": 2.5
}
```
