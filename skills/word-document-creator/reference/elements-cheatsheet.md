# Elements Cheatsheet for `word_build_document` & `word_batch_add_content`

Each element in the `elements` array is a dictionary with a `"type"` property and type-specific attributes.

## 1. Heading (`"type": "heading"`)
```json
{
  "type": "heading",
  "text": "1. Resumen Ejecutivo",
  "level": 1
}
```
* `level`: Integer between 1 and 4 (default: 1).

## 2. Paragraph (`"type": "paragraph"`)
```json
{
  "type": "paragraph",
  "text": "Texto con **negrita** y *cursiva* opcionales.",
  "alignment": "justify",
  "space_after": 6
}
```
* `alignment`: `"left"`, `"center"`, `"right"`, or `"justify"` (default: `"justify"`).
* `space_after`: Space after paragraph in points (default: 6).
* Markdown inline formatting: Supports `**bold**`, `*italic*`, and `***bold italic***` seamlessly.

## 3. Table (`"type": "table"`)
```json
{
  "type": "table",
  "headers": ["Columna 1", "Columna 2", "Columna 3"],
  "rows": [
    ["Fila 1 A", "Fila 1 B", "Fila 1 C"],
    ["Fila 2 A", "Fila 2 B", "Fila 2 C"]
  ],
  "style": "Table Grid",
  "alignment": "center"
}
```
* `headers`: Array of strings for table headers. Automatically bolded and formatted with light background shading.
* `rows`: 2D array of string cells.
* `style`: Word table style name (default: `"Table Grid"`).

## 4. Bullet List (`"type": "bullet_list"`)
```json
{
  "type": "bullet_list",
  "items": [
    "Primer elemento con detalle",
    "Segundo elemento relevante",
    "Tercer elemento"
  ]
}
```

## 5. Numbered List (`"type": "numbered_list"`)
```json
{
  "type": "numbered_list",
  "items": [
    "Paso 1: Diagnóstico inicial",
    "Paso 2: Implementación",
    "Paso 3: Validación"
  ]
}
```

## 6. Page Break (`"type": "page_break"`)
```json
{
  "type": "page_break"
}
```

## 7. Callout Box (`"type": "callout"`)
```json
{
  "type": "callout",
  "text": "Nota importante: Esta configuración requiere permisos de administrador.",
  "title": "ATENCIÓN"
}
```
* Creates a single-cell bordered table with distinct background shading and bold title to emphasize critical information.
