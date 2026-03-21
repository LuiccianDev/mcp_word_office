# Tool Usage Patterns

This guide describes safe usage patterns for MCP Word tools.

## word_create_document

Purpose: create a new `.docx` file.

Safe pattern:
1. Validate absolute path.
2. Confirm path is allowed.
3. Create once at workflow start.

## word_add_heading

Purpose: build document hierarchy.

Safe pattern:
1. Plan heading map first.
2. Keep logical order (H1 -> H2 -> H3).
3. Avoid level jumps.

## word_add_paragraph

Purpose: add content blocks.

Safe pattern:
1. Add 1-3 focused paragraphs per iteration.
2. Keep paragraphs concise.
3. Avoid unresolved placeholders.

## word_add_table

Purpose: insert structured data.

Safe pattern:
1. Prepare and validate table data first.
2. Insert table.
3. Apply table style afterwards if needed.

## word_format_text

Purpose: local text emphasis.

Safe pattern:
1. Use for targeted highlights only.
2. Apply near the end of content building.
3. Keep color/underline usage moderate.

## word_create_custom_style

Purpose: reusable styling.

Safe pattern:
1. Create only if style repeats in 3+ blocks.
2. Use clear style names.
3. Document style intent if complex.

## word_get_document_outline

Purpose: low-cost structural read.

Use this first when editing existing files.

## word_get_document_info

Purpose: metadata read without full content load.

Use this when you need title/author/stats.

## word_get_document_text

Purpose: full content read.

Use only when necessary. Prefer outline/info first.

## word_convert_to_pdf

Purpose: export final deliverable.

Safe pattern:
1. Export only if requested.
2. Validate PDF output path before exporting.
3. Run after QA completion.

## Quick Safety Checklist

- [ ] Absolute allowed path
- [ ] Correct tool for the stage
- [ ] Iterative build instead of monolithic output
- [ ] Format after structure stabilization
- [ ] Explicit final status summary
