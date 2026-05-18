---
name: office
description: Work with Microsoft Word and Excel files on Windows while preserving layout, formulas, and Office fidelity.
---

# Office

Use this skill for Microsoft Word and Excel files when formatting, layout, formulas, comments, revisions, charts, pivots, or workbook fidelity matter.

## Core rule

Do not treat `.doc`, `.docx`, `.xls`, `.xlsx`, or `.xlsm` files as plain text when the user cares about layout or formulas. Use the Office tools so Microsoft Word and Excel perform the rendering and calculation.

## Word workflow

- Use `office_word_create` when creating a new editable Word document, especially for Chinese or non-ASCII output paths.
- Start with `office_word_inspect` to understand pages, headings, tables, comments, and revisions.
- Use `office_word_export_pdf` when the user wants to verify visual layout or final rendering.
- Use `office_word_replace` for text replacements while preserving Word formatting.
- Prefer `saveAs` for edits unless the user explicitly wants to modify the original file.
- Use tracked changes when the user wants reviewable edits.

## Excel workflow

- Start with `office_excel_inspect` to understand sheets, used ranges, formulas, errors, charts, pivots, and named ranges.
- Use `office_excel_range` for focused analysis instead of dumping an entire workbook.
- Use `office_excel_recalculate` before trusting formula-dependent results.
- Use `office_excel_write` for value or formula changes; strings beginning with `=` are formulas.
- Use `office_excel_export_pdf` for layout-preserving workbook or sheet output.
- Prefer `saveAs` for edits unless the user explicitly wants to modify the original file.

## Safety

- These tools require Windows with Microsoft Word or Excel installed.
- Macros are disabled by default. Only enable macros when the user explicitly asks and accepts the risk.
- For large workbooks, inspect first, then read specific sheets and ranges.
