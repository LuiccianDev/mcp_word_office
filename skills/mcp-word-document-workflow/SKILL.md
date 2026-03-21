---
name: mcp-word-document-workflow
description: "Create and improve high-quality .docx files with the MCP Word server using a reliable 5-step workflow: brief capture, structure planning, iterative content build, formatting pass, and final QA. Use this whenever the user asks to draft, edit, standardize, or export Word documents such as reports, letters, meeting minutes, proposals, academic documents, and administrative requests."
argument-hint: "Document type, objective, audience, language/tone, and absolute output path"
user-invocable: true
disable-model-invocation: false
---

# MCP Word Document Workflow

Use this skill to produce high-quality Word documents with fewer ambiguities, less rework, and consistent output.

## When to Use

- Create new `.docx` documents: reports, letters, meeting minutes, proposals.
- Create academic and administrative documents: university reports, syllabi, formal requests.
- Improve or restructure existing Word documents.
- Apply consistent formatting (headings, tables, sections, metadata).
- Export to PDF as a final deliverable.

## Goal

Deliver a useful, well-structured, verifiable Word document through a repeatable workflow with strict path safety.

## Operational Workflow (5 Steps)

1. **Minimum Brief:** Confirm type, objective, audience, and absolute allowed path.
2. **Structure Plan:** Define document skeleton (cover, sections, special blocks).
3. **Block Construction:** Build content in small iterations (1-3 items).
4. **Enrichment:** Add tables, formatting, headers/footers, TOC.
5. **QA and Close:** Validate structure, metadata, placeholders, and final status.

## Decision Logic (Branching)

- **User asks for speed:** deliver a minimum viable version, then refine.
- **User asks for formal precision:** spend more effort on structure, styling, and QA.
- **Missing content inputs:** build a guided template and request missing fields.
- **Existing source document:** inspect with `word_get_document_outline` before full text.
- **Need consolidation:** use `word_merge_documents` and normalize formatting after merge.

## Quality Criteria

Done means:
1. The document matches objective and audience.
2. Structure is readable and consistent.
3. No obvious continuity/format issues.
4. Requested elements are present.
5. Final response clearly states what was done and what remains.

## Activation Examples

- "Create a monthly sales report with executive summary, regional table, and conclusions."
- "Take this .docx and convert it into a formal proposal with improved structure and TOC."
- "Generate meeting minutes with actions, owners, and due dates."
- "Create a university report with institutional cover, objectives, development, and bibliography."
- "Draft a formal extension request with rationale and attachments."
- "Create a 16-week syllabus with grading policy and weekly schedule."

---

## Reference Files (Important)

### Security and Paths
[reference/security.md](reference/security.md)
- Strict allowed-directory policy (`MCP_ALLOWED_DIRECTORIES`)
- Validation before create/copy/merge/export operations
- Clear failure behavior for invalid paths

### Styles and Formatting
[reference/styles-and-formatting.md](reference/styles-and-formatting.md)
- Recommended fonts, sizes, spacing
- Practical formatting patterns
- Common formatting mistakes to avoid

### Tool Usage Patterns
[reference/tools-patterns.md](reference/tools-patterns.md)
- Safe usage patterns for each tool
- When to use outline/info/text reads
- Validation checklist by operation

### Token Optimization
[reference/token-optimization.md](reference/token-optimization.md)
- Selective read strategy (outline -> info -> full text)
- Atomic content construction
- Template reuse for repetitive workloads

### Quick Templates
[templates/README.md](templates/README.md)
- Ready-to-use templates for faster output
- Placeholder replacement pattern
- Prompt starter for AI-assisted generation

---

## Document Examples

- Report: [examples/reporte.md](examples/reporte.md)
- Letter: [examples/carta.md](examples/carta.md)
- Meeting Minutes: [examples/acta.md](examples/acta.md)
- Proposal: [examples/propuesta.md](examples/propuesta.md)
- University Report: [examples/informe-universitario.md](examples/informe-universitario.md)
- Formal Request: [examples/solicitud-formal.md](examples/solicitud-formal.md)
- Syllabus: [examples/syllabus.md](examples/syllabus.md)

---

## Security Summary

- Always use absolute paths within `MCP_ALLOWED_DIRECTORIES`.
- Always validate before write operations.
- Never proceed with invalid or out-of-scope paths.

See [reference/security.md](reference/security.md) for details.
