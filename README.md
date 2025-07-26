# MCP Office Word Server

**MCP Office Word Server** es un servidor Python que implementa el Model Context Protocol (MCP) para proporcionar capacidades avanzadas de manipulación de documentos Microsoft Word (`.docx`). Este servidor permite la automatización de tareas complejas de procesamiento de documentos de manera programática.

## Tabla de Contenidos
- [Características Principales](#-características-principales)
- [Requisitos del Sistema](#-requisitos-del-sistema)
- [Instalación](#-instalación)
- [Uso Básico](#-uso-básico)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [API de Herramientas](#-api-de-herramientas)
- [Seguridad](#-seguridad)
- [Contribución](#-contribución)
- [Licencia](#-licencia)

## Características Principales

### Gestión de Documentos
- Creación de nuevos documentos con metadatos personalizados
- Listado y gestión de documentos existentes
- Fusión de múltiples documentos
- Extracción de metadatos y propiedades del documento

### Edición de Contenido
- Inserción de texto, encabezados, párrafos y saltos de página
- Gestión avanzada de tablas (creación, formato, edición)
- Inserción y manipulación de imágenes
- Búsqueda y reemplazo de texto con expresiones regulares

### Formato Avanzado
- Aplicación de estilos de párrafo y caracteres
- Formato de fuente (negrita, cursiva, subrayado, color, tamaño)
- Creación y aplicación de estilos personalizados
- Formato de tablas (bordes, alineación, sombreado)

### Protección y Seguridad
- Protección con contraseña a nivel de documento
- Restricciones de edición por secciones
- Firma digital de documentos
- Eliminación segura de metadatos

### Gestión de Referencias
- Notas al pie y notas finales
- Tablas de contenido generadas automáticamente
- Números de página y encabezados/pies de página

## Requisitos del Sistema

- Python 3.11 o superior
- Bibliotecas requeridas:
  - `python-docx>=0.8.11`
  - `msoffcrypto-tool>=4.12.0`
  - `docx2pdf>=0.1.8` (opcional, para conversión a PDF)

## Instalación

1. Clona el repositorio:
   ```bash
   git clone https://github.com/tu-usuario/mcp-office-word.git
   cd mcp-office-word
   ```

2. Crea y activa un entorno virtual (recomendado):
   ```bash
   python -m venv venv
   source venv/bin/activate  # En Windows: venv\Scripts\activate
   ```

3. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

## Uso Básico

Inicia el servidor MCP:
```bash
python mcp_word_server.py
```

### Ejemplos de Uso

#### Crear un nuevo documento
```python
create_document(
    filename="ejemplo.docx",
    title="Documento de Ejemplo",
    author="Autor",
    subject="Ejemplo de Documento"
)
```

#### Añadir contenido
```python
add_paragraph(
    filename="ejemplo.docx",
    text="Este es un párrafo de ejemplo con formato.",
    style="Heading1"
)
```

#### Buscar y reemplazar texto
```python
find_and_replace(
    filename="ejemplo.docx",
    find_text="ejemplo",
    replace_text="muestra",
    match_case=True
)
```

## Estructura del Proyecto

```
mcp-office-word/
├── mcp_word_server/
│   ├── tools/           # Herramientas MCP expuestas
│   ├── core/            # Lógica principal de manipulación de Word
│   ├── utils/           # Utilidades y funciones auxiliares
│   ├── prompts/         # Prompt templates
│   ├── validation/      # Validación de entrada
│   └── main.py          # Punto de entrada principal
├── tests/               # Pruebas unitarias
└── README.md            # Este archivo
```

## API de Herramientas

El servidor expone las siguientes categorías de herramientas:

### Documentos
- `create_document`: Crea un nuevo documento
- `list_documents`: Lista documentos en el directorio
- `merge_documents`: Combina múltiples documentos
- `protect_document`: Protege un documento con contraseña

### Contenido
- `add_paragraph`: Añade un párrafo
- `add_heading`: Añade un encabezado
- `add_table`: Crea una tabla
- `add_image`: Inserta una imagen

### Formato
- `apply_style`: Aplica un estilo
- `format_text`: Formatea texto seleccionado
- `format_table`: Formatea una tabla

### Búsqueda
- `find_text`: Busca texto en el documento
- `find_and_replace`: Busca y reemplaza texto

## Seguridad

- Todas las operaciones de protección utilizan cifrado seguro
- Las contraseñas se manejan de forma segura y nunca se almacenan en texto plano
- Validación de entrada en todas las funciones expuestas
- Manejo seguro de archivos temporales

## Contribución

Las contribuciones son bienvenidas. Por favor, lee las pautas de contribución antes de enviar pull requests.

## Licencia

Este proyecto está licenciado bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para más detalles.

---

**Desarrollado para simplificar la automatización de documentos Word en entornos empresariales y de desarrollo.**
