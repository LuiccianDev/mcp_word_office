"""
Prompt templates for MCP Office Word Server.

This module contains all the prompt templates used for generating academic documents
in the MCP Office Word Server application based on prompts_estudiantes.md.
"""

from pydantic import BaseModel


class PromptTemplate(BaseModel):
    """A class to represent a prompt template."""

    name: str
    description: str
    template: str
    parameters: list[str] | None = None
    category: str
    academic_level: str


# Secondary/High School Prompts
SECONDARY_ESSAY_BASIC = PromptTemplate(
    name="secondary_essay_basic",
    description="Basic essay template for secondary/high school students",
    template="""Crea un documento llamado "{filename}.docx" con:
- Título centrado en Arial 16pt negrita
- Tu nombre como autor
- Estructura: introducción, 3 párrafos de desarrollo, conclusión
- Interlineado 1.5
- Márgenes de 2.5 cm
- Numeración de páginas en la esquina inferior derecha
- Agregar una portada simple""",
    parameters=["filename"],
    category="essay",
    academic_level="secondary",
)

SECONDARY_ARGUMENTATIVE_ESSAY = PromptTemplate(
    name="secondary_argumentative_essay",
    description="Argumentative essay template for secondary/high school students",
    template="""Genera un ensayo argumentativo sobre {topic} en "{filename}.docx":
- Portada con título, nombre, curso, fecha
- Índice automático
- Introducción con tesis clara
- 4-5 argumentos con evidencia
- Contraargumentos y refutación
- Conclusión sólida
- Bibliografía con 5 fuentes mínimo
- Formato APA básico""",
    parameters=["topic", "filename"],
    category="essay",
    academic_level="secondary",
)

# University Level Prompts
UNIVERSITY_CRITICAL_ESSAY = PromptTemplate(
    name="university_critical_essay",
    description="Critical essay template for university students",
    template="""Desarrolla un ensayo crítico en "{filename}.docx":
- Formato APA completo (portada, headers, numeración)
- Abstract en español e inglés
- Introducción con estado del arte
- Desarrollo con análisis crítico profundo
- Uso de citas textuales y paráfrasis
- Notas al pie para aclaraciones
- Referencias en formato APA 7ma edición
- Anexos si es necesario
- {word_count} palabras""",
    parameters=["filename", "word_count"],
    category="essay",
    academic_level="university",
)

UNIVERSITY_RESEARCH_PROPOSAL = PromptTemplate(
    name="university_research_proposal",
    description="Research proposal template for university students",
    template="""Elabora una propuesta en "{filename}.docx":
- Carátula con datos institucionales
- Resumen del proyecto
- Planteamiento y justificación del problema
- Objetivos generales y específicos
- Marco teórico preliminar
- Hipótesis o preguntas de investigación
- Diseño metodológico
- Cronograma detallado
- Presupuesto estimado
- Referencias preliminares""",
    parameters=["filename"],
    category="proposal",
    academic_level="university",
)

# Laboratory Reports
SECONDARY_LAB_REPORT = PromptTemplate(
    name="secondary_lab_report",
    description="Laboratory report template for secondary/high school students",
    template="""Crea "{filename}.docx" con:
- Portada profesional (título, nombres, fecha, curso)
- Tabla de contenido automática
- Secciones: Objetivo, Materiales, Procedimiento, Resultados, Análisis, Conclusiones
- Tablas para datos experimentales
- Espacio para insertar gráficas
- Formato científico estándar
- Numeración de figuras y tablas""",
    parameters=["filename"],
    category="lab_report",
    academic_level="secondary",
)

# University Thesis Templates
UNIVERSITY_THESIS_TEMPLATE = PromptTemplate(
    name="university_thesis_template",
    description="Complete thesis template for undergraduate level",
    template="""Estructura base para tesis en "{filename}.docx":
- Portada institucional oficial
- Páginas preliminares (dedicatoria, agradecimientos, índices)
- Resumen ejecutivo bilingüe
- Capítulo 1: Planteamiento del problema
- Capítulo 2: Marco teórico
- Capítulo 3: Marco metodológico
- Capítulo 4: Análisis de resultados
- Capítulo 5: Conclusiones y recomendaciones
- Bibliografía extensa
- Anexos numerados
- Formato académico institucional""",
    parameters=["filename"],
    category="thesis",
    academic_level="university",
)

# Graduate Level Prompts
GRADUATE_RESEARCH_ARTICLE = PromptTemplate(
    name="graduate_research_article",
    description="Research article template for graduate students",
    template="""Redacta un artículo científico en "{filename}.docx":
- Formato según revista objetivo
- Título impactante y keywords
- Abstract estructurado (Objetivo, Métodos, Resultados, Conclusiones)
- Introducción con gap de conocimiento identificado
- Marco teórico con análisis crítico de literatura
- Metodología replicable y detallada
- Resultados con tablas y figuras profesionales
- Discusión comparativa con otros estudios
- Limitaciones del estudio
- Conclusiones y futuras líneas de investigación
- Referencias en formato de revista específica""",
    parameters=["filename"],
    category="research",
    academic_level="graduate",
)

GRADUATE_DOCTORAL_PROPOSAL = PromptTemplate(
    name="graduate_doctoral_proposal",
    description="Doctoral research proposal template",
    template="""Desarrolla propuesta doctoral en "{filename}.docx":
- Portada institucional completa
- Resumen ejecutivo
- Planteamiento del problema de investigación
- Revisión exhaustiva de literatura
- Gap teórico identificado
- Preguntas de investigación específicas
- Objetivos e hipótesis
- Marco teórico robusto
- Diseño metodológico innovador
- Cronograma de 3-4 años
- Contribución esperada al conocimiento
- Impacto académico y social
- Presupuesto detallado
- Referencias extensas (100+ fuentes)""",
    parameters=["filename"],
    category="proposal",
    academic_level="graduate",
)

# Business Case Studies
UNIVERSITY_CASE_STUDY = PromptTemplate(
    name="university_case_study",
    description="Business case study template for university students",
    template="""Desarrolla un estudio de caso en "{filename}.docx":
- Portada ejecutiva
- Resumen ejecutivo
- Antecedentes de la organización/situación
- Descripción del problema
- Análisis situacional (FODA)
- Alternativas de solución
- Evaluación de alternativas
- Recomendaciones específicas
- Plan de implementación
- Conclusiones
- Anexos con datos relevantes""",
    parameters=["filename"],
    category="case_study",
    academic_level="university",
)

# Format Templates
APA_FORMAT_TEMPLATE = PromptTemplate(
    name="apa_format_template",
    description="APA formatting guidelines template",
    template="""Aplica formato APA a "{filename}.docx":
- Portada con título, autor, institución, fecha
- Headers con título corto y página
- Márgenes de 1 pulgada
- Fuente Times New Roman 12pt
- Interlineado doble
- Sangría de 0.5" en primera línea
- Títulos con niveles jerárquicos APA
- Citas en texto formato (Autor, año)
- Lista de referencias ordenada alfabéticamente
- DOI cuando esté disponible""",
    parameters=["filename"],
    category="format",
    academic_level="all",
)

MLA_FORMAT_TEMPLATE = PromptTemplate(
    name="mla_format_template",
    description="MLA formatting guidelines template",
    template="""Convierte a formato MLA "{filename}.docx":
- Header con apellido y número de página
- Encabezado con datos personales
- Título centrado sin negrita
- Márgenes de 1 pulgada
- Fuente Times New Roman 12pt
- Interlineado doble
- Citas en texto (Autor página)
- Works Cited al final
- Sin portada separada""",
    parameters=["filename"],
    category="format",
    academic_level="all",
)

# Specialized Academic Documents
ACCOUNTING_REPORT_TEMPLATE = PromptTemplate(
    name="accounting_report_template",
    description="Accounting and financial report template",
    template="""Crea reporte contable en "{filename}.docx":
- Portada con información de la empresa
- Resumen ejecutivo financiero
- Estados financieros principales
- Análisis de ratios financieros
- Análisis de tendencias
- Conclusiones y recomendaciones
- Anexos con datos detallados
- Formato profesional contable""",
    parameters=["filename"],
    category="business",
    academic_level="university",
)

LEGAL_ESSAY_TEMPLATE = PromptTemplate(
    name="legal_essay_template",
    description="Legal essay template for law students",
    template="""Crea ensayo jurídico en "{filename}.docx":
- Portada con escudo institucional
- Planteamiento del problema jurídico
- Marco normativo aplicable
- Análisis de jurisprudencia relevante
- Doctrina especializada
- Análisis de casos similares
- Argumentación jurídica sólida
- Conclusiones con propuestas normativas
- Bibliografía jurídica especializada
- Anexos con textos normativos""",
    parameters=["filename"],
    category="legal",
    academic_level="university",
)

MEDICAL_CASE_STUDY = PromptTemplate(
    name="medical_case_study",
    description="Medical case study template for health science students",
    template="""Desarrolla caso clínico en "{filename}.docx":
- Datos del paciente (anonimizados)
- Motivo de consulta
- Antecedentes médicos relevantes
- Exploración física detallada
- Estudios complementarios
- Diagnóstico diferencial
- Diagnóstico final
- Plan terapéutico
- Evolución del paciente
- Discusión del caso
- Referencias médicas actualizadas""",
    parameters=["filename"],
    category="medical",
    academic_level="university",
)

# Dictionary of all available prompts
PROMPT_REGISTRY = {
    # Secondary/High School
    "secondary_essay_basic": SECONDARY_ESSAY_BASIC,
    "secondary_argumentative_essay": SECONDARY_ARGUMENTATIVE_ESSAY,
    "secondary_lab_report": SECONDARY_LAB_REPORT,
    # University Level
    "university_critical_essay": UNIVERSITY_CRITICAL_ESSAY,
    "university_research_proposal": UNIVERSITY_RESEARCH_PROPOSAL,
    "university_thesis_template": UNIVERSITY_THESIS_TEMPLATE,
    "university_case_study": UNIVERSITY_CASE_STUDY,
    # Graduate Level
    "graduate_research_article": GRADUATE_RESEARCH_ARTICLE,
    "graduate_doctoral_proposal": GRADUATE_DOCTORAL_PROPOSAL,
    # Format Templates
    "apa_format_template": APA_FORMAT_TEMPLATE,
    "mla_format_template": MLA_FORMAT_TEMPLATE,
    # Specialized Academic Documents
    "accounting_report_template": ACCOUNTING_REPORT_TEMPLATE,
    "legal_essay_template": LEGAL_ESSAY_TEMPLATE,
    "medical_case_study": MEDICAL_CASE_STUDY,
}
