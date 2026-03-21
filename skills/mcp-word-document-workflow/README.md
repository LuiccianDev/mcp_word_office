# MCP Word Document Workflow - Navigation Guide

This skill provides a practical system for creating high-quality Word documents with MCP.

## Content Structure

```text
mcp-word-document-workflow/
├── SKILL.md
├── README.md
├── reference/
│   ├── security.md
│   ├── styles-and-formatting.md
│   ├── tools-patterns.md
│   └── token-optimization.md
├── templates/
│   ├── README.md
│   ├── reporte-template.md
│   ├── carta-template.md
│   ├── acta-template.md
│   ├── propuesta-template.md
│   ├── informe-universitario-template.md
│   ├── solicitud-formal-template.md
│   └── syllabus-template.md
├── examples/
│   ├── reporte.md
│   ├── carta.md
│   ├── acta.md
│   ├── propuesta.md
│   ├── informe-universitario.md
│   ├── solicitud-formal.md
│   └── syllabus.md
└── evals/
    └── evals.json
```

## Skill Name Decision

Keep the name `mcp-word-document-workflow`.

Why:
- Stable triggering behavior.
- Broad enough for business, academic, and administrative use cases.
- No migration cost for existing prompts/references.

## Quick Start

1. Read [SKILL.md](SKILL.md).
2. Pick a document type in [examples](examples).
3. Use technical references in [reference](reference).
4. Start fast from [templates/README.md](templates/README.md).
5. Validate quality with [evals/evals.json](evals/evals.json).

## Document Types

- Report: [examples/reporte.md](examples/reporte.md)
- Letter: [examples/carta.md](examples/carta.md)
- Meeting minutes: [examples/acta.md](examples/acta.md)
- Proposal: [examples/propuesta.md](examples/propuesta.md)
- University report: [examples/informe-universitario.md](examples/informe-universitario.md)
- Formal request: [examples/solicitud-formal.md](examples/solicitud-formal.md)
- Syllabus: [examples/syllabus.md](examples/syllabus.md)

## Security Golden Rules

- Use absolute paths only.
- Keep all output under `MCP_ALLOWED_DIRECTORIES`.
- Validate path before every write operation.
- Fail fast and request a valid path if needed.

## Final Checklist

- [ ] Structure complete for selected document type
- [ ] No unresolved placeholders (`TODO`, `[PENDING]`, `lorem`)
- [ ] Basic metadata set (title/author/date when needed)
- [ ] Consistent formatting
- [ ] Output saved in allowed absolute path
- [ ] Clear completion summary sent to user

## Changelog

- **v1.0** - Initial modular restructure (`SKILL` + `reference` + `examples`).
- **v1.1** - Added academic/administrative examples and templates.
- **v1.2** - Quality pass and evaluation scaffolding.
- **v1.3** - Full English localization across the skill package.
