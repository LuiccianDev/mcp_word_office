<div align="center">
  <h1> MCP Office Word Server</h1>
  <p>
    <em>Potente servidor para la manipulación programática de documentos Word (.docx) mediante MCP</em>
  </p>
  
  [![Python Version](https://img.shields.io/badge/python-3.13+-blue.svg)](https://www.python.org/)
  [![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
  [![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)

</div>

**MCP Office Word Server** es un servidor Python que implementa el Model Context Protocol (MCP) para proporcionar capacidades avanzadas de manipulación de documentos Microsoft Word (`.docx`). Este servidor permite la automatización de tareas complejas de procesamiento de documentos de manera programática.

> 💡 **Nota**: Este proyecto está diseñado para ser utilizado con clientes compatibles con MCP como Claude, permitiendo la manipulación de documentos Word mediante instrucciones en lenguaje natural.

## 📑 Tabla de Contenidos

- [✨ Características Principales](#-características-principales)
- [🖥️ Requisitos del Sistema](#️-requisitos-del-sistema)
- [⚙️ Instalación](#️-instalación)
- [🗂️ Estructura del Proyecto](#️-estructura-del-proyecto)
- [🔧 API de Herramientas](#-api-de-herramientas)
- [� Seguridad](#-seguridad)
- [🤝 Contribución](#-contribución)
- [📜 Licencia](#-licencia)

## ✨ Características Principales

### 📂 Gestión de Documentos

- 📝 **Creación de documentos** con metadatos personalizados
- 📋 **Gestión de archivos** existentes y listado de documentos
- 🔄 **Fusión** de múltiples documentos en uno solo
- 🔍 **Extracción** de metadatos y propiedades del documento

### ✏️ Edición de Contenido

- 📝 **Inserción avanzada** de texto, encabezados y párrafos
- 📊 **Gestión de tablas** con múltiples opciones de formato
- 🖼️ **Manipulación de imágenes** con diferentes estilos
- 🔍 **Búsqueda y reemplazo** con soporte para expresiones regulares

### 🎨 Formato Avanzado

- 🖍️ **Estilos personalizables** para párrafos y caracteres
- 🔠 **Opciones de fuente** completas (familia, tamaño, color, efectos)
- 🎭 **Temas y estilos** predefinidos y personalizables
- 📐 **Formateo de tablas** con bordes, colores y alineación

### 🔒 Protección y Seguridad

- 🔐 **Protección** con contraseña a nivel de documento
- 🛡️ **Restricciones** de edición por secciones
- 📜 **Firma digital** para autenticidad del documento
- 🗑️ **Eliminación segura** de metadatos sensibles

### 📚 Gestión de Referencias

- 📌 **Notas al pie** y notas finales personalizables
- 📑 **Tablas de contenido** generadas automáticamente
- 🔢 **Numeración** de páginas y secciones
- 📖 **Índices** y referencias cruzadas

## 🖥️ Requisitos del Sistema

### 📋 Requisitos Mínimos

- **Python**: 3.13 o superior
- **Sistema Operativo**: Windows, macOS o Linux
- **Memoria RAM**: 2 GB mínimo (4 GB recomendado)
- **Espacio en disco**: 100 MB libres

## ⚙️ Instalación

### Requisitos Previos

- Python 3.13 o superior
- [uv](https://github.com/astral-sh/uv) - Instalador y gestor de entornos ultra rápido

### 1. Clonar el repositorio

```bash
git clone https://github.com/LuiccianDev/mcp_office_word.git
cd mcp_office_word
```

### 2. Configurar entorno virtual con uv

uv es el gestor de paquetes recomendado para este proyecto. Crea y activa automáticamente un entorno virtual:

```bash
# Crear y activar entorno virtual
uv venv

# En Windows:
.venv\Scripts\activate

# En macOS/Linux:
source .venv/bin/activate
```

### 3. Instalar dependencias

Instala las dependencias del proyecto usando uv:

```bash
uv pip install -e ".[dev]"
```

> ℹ️ El comando anterior instalará tanto las dependencias principales como las de desarrollo.

## 🔌 Configuración para Clientes MCP

### Enlaces Rápidos

- [Repositorio](https://github.com/LuiccianDev/mcp_word_office)
- [Documentación](https://github.com/LuiccianDev/mcp_word_office/blob/master/README.md)
- [Reportar un Problema](https://github.com/LuiccianDev/mcp_word_office/issues)
- [Registro de Cambios](https://github.com/LuiccianDev/mcp_word_office/blob/master/CHANGELOG.md)

### Configuración Básica

Para integrar el servidor MCP Office Word con clientes compatibles como Claude, sigue estos pasos:

1. **Inicia el servidor** siguiendo las instrucciones en la sección
2. **Configura tu cliente MCP** con los siguientes parámetros:

```json
{
  "mcp-word-office": {
    "command": "python",
    "args": ["Users/path/to/mcp_server.py"],
    "env": {
      "MCP_ALLOWED_DIRECTORIES": "Users/path/to/your/documents"
    }
  }
}
```

### 🔧 Variables de Entorno Clave

| Variable                  | Descripción                                                  | Ejemplo                                          |
| ------------------------- | ------------------------------------------------------------ | ------------------------------------------------ |
| `MCP_ALLOWED_DIRECTORIES` | Directorios accesibles por el servidor (separados por comas) | `"\Users\Usuario\Documentos,.Proyectos"` |

### 🔒 Consideraciones de Seguridad

- 🔐 **Directorios Permitidos**: Limita los directorios accesibles a solo los necesarios.
- 🛡️ **Entorno Virtual**: Siempre usa un entorno virtual para aislar las dependencias.
- 🔄 **Actualizaciones**: Mantén el servidor actualizado con las últimas correcciones de seguridad.
- 👥 **Permisos**: Asegúrate de que los permisos de archivo sean los adecuados.

## 🗂️ Estructura del Proyecto

```text
mcp-office-word/
│    └── 📁 word_mcp/            # Paquete principal del servidor
│        ├── 📁 core/            # Lógica principal de manipulación de Word
│        ├── 📁 tools/           # Herramientas MCP expuestas
│        ├── 📁 utils/           # Utilidades y funciones auxiliares
│        ├── 📁 prompts/         # Plantillas de prompts para MCP
│        ├── 📁 validation/      # Validación de entrada
│        └── main.py             # Punto de entrada principal
│
├── 📁 tests/                    # Pruebas unitarias
├── 📄 README.md                 # Este archivo
└── 📄 pyproject.toml            # Configuración del proyecto
```

### 📋 Descripción de Directorios

- **` word_mcp/`**: Contiene todo el código fuente del servidor.

  - **`core/`**: Lógica central para la manipulación de documentos Word.
  - **`tools/`**: Implementación de las herramientas expuestas a través de MCP.
  - **`utils/`**: Funciones auxiliares compartidas.
  - **`prompts/`**: Plantillas para generar instrucciones para el modelo de lenguaje.
  - **`validation/`**: Validación de entradas y parámetros.

- **`tests/`**: Pruebas unitarias y de integración para garantizar el correcto funcionamiento.

## 🔧 API de Herramientas

El servidor MCP Office Word expone un conjunto completo de herramientas organizadas en categorías lógicas para facilitar la manipulación de documentos Word.

### 📄 Documentos

| Herramienta        | Descripción                                | Parámetros                               |
| ------------------ | ------------------------------------------ | ---------------------------------------- |
| `create_document`  | Crea un nuevo documento Word               | `filename`, `title`, `author`, `subject` |
| `list_documents`   | Lista documentos en directorios permitidos | `directory` (opcional)                   |
| `merge_documents`  | Combina múltiples documentos               | `target_filename`, `source_filenames`    |
| `protect_document` | Protege un documento con contraseña        | `filename`, `password`                   |

### 📝 Contenido

| Herramienta     | Descripción               | Parámetros                         |
| --------------- | ------------------------- | ---------------------------------- |
| `add_paragraph` | Añade un párrafo de texto | `filename`, `text`, `style`        |
| `add_heading`   | Añade un encabezado       | `filename`, `text`, `level`        |
| `add_table`     | Crea una tabla            | `filename`, `rows`, `cols`, `data` |
| `add_image`     | Inserta una imagen        | `filename`, `image_path`, `width`  |

### 🎨 Formato

| Herramienta    | Descripción                    | Parámetros                           |
| -------------- | ------------------------------ | ------------------------------------ |
| `apply_style`  | Aplica un estilo a un elemento | `filename`, `element_id`, `style`    |
| `format_text`  | Formatea un rango de texto     | `filename`, `start`, `end`, `format` |
| `format_table` | Formatea una tabla             | `filename`, `table_index`, `style`   |

### 🔍 Búsqueda

| Herramienta        | Descripción                 | Parámetros                              |
| ------------------ | --------------------------- | --------------------------------------- |
| `find_text`        | Busca texto en el documento | `filename`, `search_text`               |
| `find_and_replace` | Busca y reemplaza texto     | `filename`, `find_text`, `replace_text` |

### 📚 Referencias

| Herramienta    | Descripción                     |
| -------------- | ------------------------------- |
| `add_footnote` | Añade una nota al pie           |
| `add_endnote`  | Añade una nota final            |
| `update_toc`   | Actualiza la tabla de contenido |

> ℹ️ Para una documentación detallada de cada herramienta, consulta el archivo [TOOLS.md](TOOLS.md).

## 🔒 Seguridad

### Consideraciones de Seguridad

1. **Validación de Entrada**

   - Todas las funciones realizan validación estricta de parámetros
   - Se utilizan tipos de datos fuertemente tipados
   - Se aplica sanitización de rutas de archivo

2. **Seguridad de Archivos**

   - Uso de `MCP_ALLOWED_DIRECTORIES` para restringir acceso
   - Manejo seguro de archivos temporales
   - Validación de tipos MIME para archivos subidos

3. **Buenas Prácticas**
   - Código revisado con `mypy` para seguridad de tipos
   - Análisis estático con `ruff`
   - Pruebas unitarias para casos de seguridad

## 🤝 Contribución

Las contribuciones son bienvenidas. Por favor, lee las pautas de contribución antes de enviar pull requests.

## 📜 Licencia

Este proyecto está licenciado bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para más detalles.

---

**Desarrollado para simplificar la automatización de documentos Word en entornos empresariales y de desarrollo.**
