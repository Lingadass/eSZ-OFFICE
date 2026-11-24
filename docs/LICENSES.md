# Third-Party Licenses — eSZ OFFICE

This document tracks the open-source licenses of key third-party rendering, document processing, and font assets bundled in **eSZ OFFICE**.

## Document Rendering & OOXML Libraries

| Package | Version | License | Author / Source | Notes |
| --- | --- | --- | --- | --- |
| `docx-preview` | `0.4.1` | **Apache-2.0** | Volodymyr Baydalka — https://github.com/VolodymyrBaydalka/docxjs | Used in eSZ OFFICE View Mode to render original `.docx` files directly from stored OOXML bytes without mutating the source file. Verified from `node_modules/docx-preview/package.json` and `dist/docx-preview.d.ts`. |
| `jszip` | `^3.10.2` | **MIT OR GPL-3.0-or-later** | Stuart Knightley — https://github.com/Stuk/jszip | Used for reading and writing OOXML `.docx`, `.xlsx`, `.pptx`, and `.zip` archive containers locally on-device. |
| `docx` | `^9.15.4` | **MIT** | Dolan Miu — https://github.com/dolanmiu/docx | Used for generating and exporting `.docx` documents. |
| `mammoth` | `^1.13.0` | **BSD-2-Clause** | Michael Williamson — https://github.com/mwilliamson/mammoth.js | Legacy fallback `.docx` HTML converter (being superseded by custom OOXML importer in Step 3). |
| `pdfjs-dist` | `^6.5.438` | **Apache-2.0** | Mozilla Foundation — https://github.com/mozilla/pdf.js | Local PDF rendering engine. |
| `pdf-lib` | `^1.17.1` | **MIT** | Andrew Dillon — https://github.com/Hopding/pdf-lib | Local PDF creation, annotation, and page manipulation. |

## Bundled Offline Fonts

All offline document fonts in `/public/fonts/*.woff2` are self-hosted and precached by the service worker under the **SIL Open Font License 1.1 (OFL-1.1)**:

| Font Family | Metric Substitute For | Version / Source | License | Files |
| --- | --- | --- | --- | --- |
| **Liberation Serif** | Times New Roman | `2.1.5` — Red Hat, Inc. / Google Corp. (`https://github.com/liberationfonts/liberation-fonts`) | **SIL OFL 1.1** | `liberation-serif-{400,700}-{normal,italic}.woff2` |
| **Liberation Sans** | Arial | `2.1.5` — Red Hat, Inc. / Google Corp. (`https://github.com/liberationfonts/liberation-fonts`) | **SIL OFL 1.1** | `liberation-sans-{400,700}-{normal,italic}.woff2` |
| **Liberation Mono** | Courier New | `2.1.5` — Red Hat, Inc. / Google Corp. (`https://github.com/liberationfonts/liberation-fonts`) | **SIL OFL 1.1** | `liberation-mono-{400,700}-{normal,italic}.woff2` |
| **Carlito** | Calibri | `5.3.0` (`@fontsource/carlito`, upstream Google Crosextra / Łukasz Dziedzic) | **SIL OFL 1.1** | `carlito-{400,700}-{normal,italic}.woff2` |
| **Caladea** | Cambria | `5.3.0` (`@fontsource/caladea`, upstream Google Crosextra / Huerta Tipográfica) | **SIL OFL 1.1** | `caladea-{400,700}-{normal,italic}.woff2` |

