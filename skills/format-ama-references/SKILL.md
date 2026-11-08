---
name: format-ama-references
description: Format, clean, verify, and rebuild academic reference lists in Microsoft Word according to the current American Marketing Association journal reference style. Use for AMA-style bibliographies for Journal of Marketing, Journal of Marketing Research, Journal of Public Policy & Marketing, Journal of International Marketing, Journal of Interactive Marketing, or related marketing manuscripts; for checking author names and bibliographic metadata; for using an optional reference-check/template document; and for fixing title case, italics, ordering, page ranges, hanging indents, spacing, or inconsistent Word formatting.
---

# Format AMA References

Create a publication-ready Word reference list while preserving every source and distinguishing factual verification from mechanical formatting.

## Required workflow

1. Identify the source reference-list document, the desired output, and any optional reference-check or template document.
2. Use the `documents` skill for all `.docx` work and follow its render-and-verify workflow.
3. Open the current official AMA guidance before changing references:
   - [AMA journal submission guidelines](https://www.ama.org/submission-guidelines-american-marketing-association-journals/)
   - [AMA reference-style examples](https://www.ama.org/american-marketing-association-journals-reference-style-examples/)
4. Read [references/ama-formatting-checklist.md](references/ama-formatting-checklist.md). Treat it as an operational aid; the current official AMA pages control when they differ.
5. Inventory and reconstruct the complete source list before editing.
6. Verify uncertain bibliographic facts against authoritative sources.
7. Normalize the citation text and encode each finished entry as structured segments.
8. Generate the Word file with [scripts/format_reference_docx.py](scripts/format_reference_docx.py).
9. Render the result and inspect every page. Iterate until both content and layout checks pass.
10. Save or update the user-facing file in its required persistent location and report material corrections or unresolved items.

## Inventory the source without omissions

- Extract references from body paragraphs, tables, headers, footers, text boxes, footnotes, and endnotes when present. Inspect the underlying DOCX XML if the ordinary paragraph API misses visible text.
- Rejoin entries that were accidentally split across paragraphs or table cells.
- Record a stable count and a short identifier for every source before rewriting anything.
- Compare the final list with this inventory. Do not silently drop, merge, or duplicate references.
- Use a supplied “Reference Check,” prior bibliography, or template as secondary evidence for formatting and missing metadata. Do not assume it supersedes current AMA guidance or the publisher record.

## Verify bibliographic facts

Research whenever an author name, publication year, title, venue, volume, issue, page range, DOI, URL, conference, or publication status is uncertain. Prefer sources in this order:

1. Official journal or publisher article page.
2. DOI landing page or official Crossref metadata.
3. Official university, institutional repository, conference, or proceedings record.
4. Author CV or personal publication page for working papers and unpublished material.
5. Indexing services only as corroboration.

Use the version the manuscript actually cites. Distinguish issue year, online-first year, working-paper revision year, and conference year. Never guess a full first name, page range, issue, DOI, or publication status. If authoritative sources conflict, preserve the defensible version and briefly flag the conflict for the user.

If the user prohibits web research or access is unavailable, format from the supplied evidence and mark uncertain fields instead of inventing them.

## Normalize the reference text

- Apply the current AMA templates for the source type; do not force journal-article syntax onto books, chapters, proceedings, working papers, dissertations, news, or websites.
- Use full author names when reliable source metadata supplies them. Preserve initials when the full name cannot be verified.
- Apply AMA title-style capitalization and punctuation consistently.
- Italicize publication containers such as journal and book titles as required by the current examples; do not italicize the entire entry.
- Abbreviate inclusive page ranges according to AMA/Chicago conventions and use en dashes.
- Alphabetize by the first author’s surname and apply the current AMA rule for repeated authors and same-year suffixes.
- Keep each complete reference in one Word paragraph.

## Build the DOCX

Create a UTF-8 JSON file containing the fully normalized entries in final order. Each entry can be a plain string or an object with styled segments:

```json
[
  {
    "segments": [
      {"text": "Author, Full Name (2025), “Article Title,” ", "italic": false},
      {"text": "Journal Title", "italic": true},
      {"text": ", 89 (1), 1–18.", "italic": false}
    ]
  }
]
```

Run:

```bash
"${CODEX_PRIMARY_RUNTIME_PYTHON:-python3}" scripts/format_reference_docx.py \
  --input source.docx \
  --references finalized_references.json \
  --output formatted_references.docx
```

The formatter rebuilds the document body with a centered “References” heading, Times New Roman 12-point type, one-inch margins, one reference per paragraph, single spacing, and a 0.3-inch hanging indent. Use `--spacing double` only when the manuscript or user explicitly requires double spacing.

## Render and quality-check

Follow the `documents` skill to render the DOCX to page images and inspect every page at readable resolution.

Content gates:

- Final reference count equals the reconstructed source count, except for explicitly documented additions or removals.
- No duplicates, fragments, placeholder text, or unresolved accidental line breaks remain.
- Alphabetization, author names, years, titles, venues, volumes, issues, and pages are internally consistent.
- Every factual change is supported by an authoritative record or reported as unresolved.

Layout gates:

- “References” is centered and visually separated from the first entry.
- All body text is Times New Roman, 12 point.
- Each entry has a true 0.3-inch hanging indent and consistent spacing.
- Only required publication containers are italicized.
- No reference is clipped, hidden, overlapping, or stranded by an avoidable page-break problem.
- Headers, footers, tables, and legacy formatting from the source do not survive unless the user requested them.

Deliver the formatted `.docx` and a concise note listing the number of references, material metadata corrections, unresolved uncertainties, and any intentional deviations from AMA style.
