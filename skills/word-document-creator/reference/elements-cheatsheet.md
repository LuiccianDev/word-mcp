# Elements Cheatsheet for `build_document` & `batch_add_content`

Each element in the `elements` array is a dictionary with a `"type"` property and type-specific attributes.

## 1. Heading (`"type": "heading"`)
```json
{
  "type": "heading",
  "text": "1. Resumen Ejecutivo",
  "level": 1
}
```
* `level`: Integer between 1 and 9 (default: 1).

## 2. Paragraph (`"type": "paragraph"`)
```json
{
  "type": "paragraph",
  "text": "Texto con **negrita** y *cursiva* opcionales.",
  "alignment": "justify",
  "space_after": 6
}
```
* `alignment`: `"left"`, `"center"`, `"right"`, or `"justify"` (Default: `left`).
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
* `headers`: Array of strings for table headers. Headers are bold only; no background shading is applied.
* `rows`: 2D array of string cells.
* `style`: Word table style name (default: `"Table Grid"`). Table alignment is always centered and the `alignment` element key is ignored.

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
  "title": "ATENCIÓN",
  "variant": "warning"
}
```
* When `variant` is omitted and `title` is provided, `title` is rendered as a bold prefix (`{title}: `); it renders an indented italic paragraph.
* When `variant` (`"note"`, `"tip"`, `"important"`, `"warning"`, or `"caution"`) is provided, it renders a shaded 1x1 callout table with a colored left accent border.

## 8. Image (`"type": "image"`)
```json
{
  "type": "image",
  "path": "/path/to/diagrama.png",
  "width": 4.5,
  "caption": "Figura 1: Arquitectura del sistema",
  "alignment": "center"
}
```
* `path`: Must reside inside `WORD_MCP_DIRS`.
* `width`: Optional width in inches (positive float).

## 9. Table of Contents (`"type": "toc"`), Title Block (`"type": "title_block"`), & References (`"type": "references"`)
```json
{"type": "title_block", "title": "Informe Técnico", "subtitle": "Fase 1", "author": "Equipo Ingeniería", "institution": "División Técnica", "date": "2026-11-18"}
{"type": "toc", "title": "Tabla de Contenido", "max_level": 3, "page_break": true}
{"type": "references", "title": "Referencias", "items": ["Autor, A. (2026). Título de la obra. Editorial."]}
```
