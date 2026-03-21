# Token Optimization

Reduce token usage while preserving quality and control.

## Core Strategy

Use a balanced approach:
- prioritize clarity and correctness
- avoid unnecessary reads
- build content iteratively

## Selective Reading

Use the lowest-cost read tool first:
1. `word_get_document_outline` for structure
2. `word_get_document_info` for metadata
3. `word_get_document_text` only when full content is required

## Atomic Construction

- Build in small batches (1-3 items).
- Validate locally after each batch.
- Apply formatting after core content is stable.

## Reuse Pattern

For repetitive documents:
1. Create a base template once.
2. Duplicate with `word_copy_document`.
3. Edit only variable fields.

This can reduce cost significantly for 3+ similar files.

## Efficiency Checklist

- [ ] Avoid full-text reads unless required
- [ ] Use outline/info before text
- [ ] Build iteratively
- [ ] Reuse templates for repeated workflows
- [ ] Keep final response concise and explicit
