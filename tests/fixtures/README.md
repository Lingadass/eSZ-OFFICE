# Document Fidelity Test Fixtures

This directory holds `.docx` test fixtures used by Vitest unit tests and Playwright visual/layout checks.

## Fixtures

- **`calibri-sample.docx`** *(generated locally)*: A clean Microsoft Word OOXML document using `Calibri` 11pt `docDefaults`/`Normal` styles, Heading 1 and Heading 2 styles, bulleted and numbered lists (`numbering.xml`), a 4-column table with explicit `w:tblGrid` widths, cell borders and header shading, a centred embedded PNG graphic, and a header/footer with `PAGE` and `NUMPAGES` fields across 2 pages.
- **`AIML.docx`** *(user-supplied — not checked in / not synthetically generated)*: Place your real `AIML.docx` file at `tests/fixtures/AIML.docx`. All `AIML.docx`-specific Vitest and Playwright checks detect the file automatically and are cleanly marked skipped/pending until it is present.
- **`VELAVAN_FRONT_PAGE_FOR_INTERSHIP.docx`** *(user-supplied original Word file — 2 pages in Word & LibreOffice)*: Place in `tests/fixtures/VELAVAN_FRONT_PAGE_FOR_INTERSHIP.docx`. Automatically enables the Step A–D LibreOffice comparison, lossless save round-trip, and PDF-export fidelity tests.
- **`VELAVAN_FRONT_PAGE_FOR.docx`** & **`VELAVAN_FRONT_PAGE_FOR.pdf`** *(user-supplied baseline eSZ re-save / export samples for regression comparison)*: Place alongside the original in `tests/fixtures/`.

