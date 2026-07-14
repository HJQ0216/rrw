---
name: mywritting
description: Reproduce and audit the writing, typography, page layout, heading hierarchy, numbering, figures, tables, captions, equations, Chinese/English text, table of contents, headers/footers, citations, and references used in the user's steel-structure master's thesis and published journal paper. Use when Codex is asked to write, format, revise, typeset, or check a Chinese thesis, graduation dissertation, academic paper, Word document, chapter, abstract, figure/table, bibliography, 论文排版, 毕业论文格式, 小论文格式, or to make new documents match 吴仁彬论文/师兄论文/mywritting格式.
---

# mywritting

Apply the measured house style from the user's two source documents. Treat the master's thesis and the journal article as separate modes; never mix their page structures.

## Select the mode

- Use **thesis mode** for dissertations, graduation theses, chapters, Word manuscripts, contents pages, appendices, acknowledgements, and degree-paper deliverables. Read [references/thesis-format.md](references/thesis-format.md) completely.
- Use **journal mode** for short papers, submissions, accepted manuscripts, or layouts matching *Science Technology and Engineering*. Read [references/journal-format.md](references/journal-format.md) completely.
- If the user does not identify the mode, infer it from the deliverable. Ask only when both are genuinely plausible and the choice would materially change the output.
- When formatting an existing DOCX, also use the `docx` skill. When reading or checking a PDF, also use the `pdf` skill.

## Source authority

Use this priority order:

1. A current school, journal, or supervisor template supplied by the user.
2. Explicit formatting instructions in the current request.
3. The measured rules in this skill.

Distinguish measured facts from editorial normalization. The references mark uncertain or visually inferred values. Do not present inferred values as official institutional requirements.

## Workflow

1. Identify the mode and output file type.
2. Read the relevant format reference in full.
3. Inventory the manuscript: front matter, headings, body, equations, figures, tables, citations, references, appendices, and bilingual blocks.
4. Build named styles before formatting individual paragraphs. Avoid manual per-paragraph formatting when a reusable style is possible.
5. Apply page setup, section breaks, numbering, headers/footers, and columns before placing figures and tables.
6. Apply Chinese and Latin fonts separately. Keep variables italic and units upright.
7. Generate captions and cross-references with automatic fields whenever the file format supports them.
8. Generate the thesis table of contents from heading levels; never type page numbers manually. Do not add a contents page in journal mode.
9. Format references only after citation order and bibliographic fields are stable.
10. Run the QA checklist in [references/qa-checklist.md](references/qa-checklist.md) before delivery.

## Non-negotiable rules

- Preserve the user's technical content, equations, symbols, data, figure meaning, and citation mapping while changing format.
- Do not invent missing bibliography fields, English translations, DOI values, author affiliations, funding numbers, or figure data. Mark them for confirmation.
- Do not stretch images or use screenshots when an original vector/high-resolution asset exists.
- Keep each figure/table close to its first substantive mention and prevent a title from being stranded at a page or column bottom.
- Use nonbreaking spaces or equivalent protection for values with units, figure/table labels, and initials where supported.
- Recalculate the table of contents, figure/table lists, page numbers, cross-references, and fields after pagination changes.
- For a new document, deliver an editable DOCX unless the user requests another format; export a PDF only after the DOCX passes visual QA.

## Minimum delivery report

State the selected mode, output path, format elements applied, validation performed, and any unresolved items that require the user's or supervisor's confirmation.
