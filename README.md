# MCP Office Word Server

**MCP Office Word Server** es un servidor Python que permite la creación, edición, protección y manipulación avanzada de documentos Microsoft Word (`.docx`) a través del Model Context Protocol (MCP). Este proyecto está diseñado para exponer operaciones de alto nivel sobre documentos Word de manera programática y automatizada.

## Características principales

- **Creación y gestión de documentos Word**
  - Crear nuevos documentos con metadatos personalizados (título, autor).
  - Listar, copiar y fusionar documentos existentes.
  - Obtener propiedades, estructura y texto completo de documentos.

- **Edición de contenido**
  - Añadir encabezados, párrafos, tablas, imágenes y saltos de página.
  - Eliminar párrafos y realizar búsqueda y reemplazo de texto.
  - Añadir tablas de contenido generadas automáticamente.

- **Formato avanzado**
  - Aplicar formato a rangos de texto (negrita, cursiva, subrayado, color, fuente, tamaño).
  - Crear estilos personalizados y aplicarlos.
  - Formatear tablas (bordes, sombreado, filas de encabezado).

- **Protección y seguridad**
  - Proteger documentos con contraseña y restringir edición por secciones.
  - Añadir y verificar firmas digitales.
  - Eliminar protección y desencriptar documentos.

- **Gestión de notas al pie y notas finales**
  - Añadir notas al pie y notas finales a párrafos específicos.
  - Convertir notas al pie en notas finales.
  - Personalizar el formato y numeración de las notas.

- **Herramientas extendidas**
  - Extraer texto de párrafos específicos.
  - Buscar texto en todo el documento (con opciones de mayúsculas/minúsculas y palabra completa).
  - Convertir documentos Word a PDF (requiere LibreOffice o Microsoft Word).

## ¿Qué es MCP (Model Context Protocol)?

MCP es un protocolo que permite exponer funciones de manipulación de documentos como "herramientas" accesibles desde clientes compatibles. Cada función del servidor puede ser llamada de forma remota, permitiendo la integración con asistentes, bots, automatizaciones y otros sistemas.

## Estructura del repositorio

- `mcp_word_server/tools/`  
  Herramientas de alto nivel expuestas por el servidor (creación, edición, formato, protección, notas, etc.).

- `mcp_word_server/core/`  
  Funcionalidad interna y utilidades para manipulación de Word (estilos, tablas, protección, notas).

- `mcp_word_server/utils/`  
  Utilidades generales para manejo de archivos, extracción de texto, búsqueda, etc.

- `mcp_word_server/main.py`  
  Punto de entrada principal del servidor MCP.

- `mcp_word_server.py`  
  Script ejecutable para iniciar el servidor.

## Ejemplo de uso

Puedes iniciar el servidor ejecutando:

```bash
python mcp_word_server.py
```

Luego, desde un cliente MCP compatible, puedes invocar herramientas como:

- Crear un documento:
  - `create_document(filename="nuevo.docx", title="Mi Documento", author="Autor")`
- Añadir un párrafo:
  - `add_paragraph(filename="nuevo.docx", text="Este es un párrafo de ejemplo.")`
- Proteger un documento:
  - `protect_document(filename="nuevo.docx", password="segura123")`
- Buscar texto:
  - `find_text_in_document(filename="nuevo.docx", text_to_find="ejemplo")`
- Convertir a PDF:
  - `convert_to_pdf(filename="nuevo.docx")`

## Requisitos

- Python 3.8+
- Paquetes: `python-docx`, `msoffcrypto-tool`, `docx2pdf` (opcional), `libreoffice` (para conversión a PDF en Linux/Mac)
- Microsoft Word (solo si se usa `docx2pdf` en Windows)

## Extensibilidad

El servidor está diseñado para ser fácilmente ampliable. Puedes agregar nuevas herramientas en la carpeta `tools/` y exponerlas a través del protocolo MCP.

## Seguridad

- El servidor no ejecuta código arbitrario ni accede a recursos externos sin autorización.
- Las operaciones de protección y firma digital usan hashing seguro y cifrado real de archivos.
- El manejo de errores y permisos de archivos es robusto para evitar corrupción de documentos.

## Licencia

MIT License.

---

**Desarrollado para automatización y manipulación avanzada de documentos Word en entornos empresariales, educativos y de investigación.**
