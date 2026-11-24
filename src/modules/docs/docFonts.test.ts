import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import {describe, expect, it} from 'vitest';
import {
  BUNDLED_FONT_FACES,
  canonicalizeWordFontFamily,
  DOCUMENT_FONT_PRESETS,
  FONT_SUBSTITUTION_MAP,
  getSubstituteFontFamily,
  isAppShellUiFont,
  resolveRenderFontStack,
  rewriteCssFontFamiliesForRender,
} from './docFonts';
import {htmlToDocxBytes} from './docsPure';
import {DEFAULT_DOC_LAYOUT, DOC_THEMES} from './docsTypes';
import {
  extractDocxDefaultTypography,
  parseDocDefaultsFromStyles,
} from './docxPreviewUtils';

const ROOT_DIR = process.cwd();
const PUBLIC_DIR = path.resolve(ROOT_DIR, 'public');
const CALIBRI_FIXTURE_PATH = path.resolve(ROOT_DIR, 'tests/fixtures/calibri-sample.docx');

describe('Step 2: Bundled SIL OFL offline fonts & substitution map', () => {
  it('bundles all 20 metric-compatible SIL OFL .woff2 font files in /public/fonts', () => {
    expect(BUNDLED_FONT_FACES).toHaveLength(20);
    for (const face of BUNDLED_FONT_FACES) {
      const diskPath = path.join(PUBLIC_DIR, face.url.replace(/^\//, ''));
      expect(fs.existsSync(diskPath), `Missing font asset ${face.url}`).toBe(true);
      const buf = fs.readFileSync(diskPath);
      expect(buf.length).toBeGreaterThan(10_000);
      expect(buf.subarray(0, 4).toString('ascii')).toBe('wOF2');
    }
  });

  it('declares @font-face rules in src/index.css for every bundled .woff2 asset', () => {
    const css = fs.readFileSync(path.resolve(ROOT_DIR, 'src/index.css'), 'utf8');
    for (const face of BUNDLED_FONT_FACES) {
      expect(css).toContain(`url("${face.url}") format("woff2")`);
      expect(css).toContain(`font-family: "${face.family}"`);
    }
  });

  it('implements the exact Word -> SIL OFL font substitution map for rendering only', () => {
    expect(FONT_SUBSTITUTION_MAP).toEqual({
      'Times New Roman': 'Liberation Serif',
      Arial: 'Liberation Sans',
      Calibri: 'Carlito',
      Cambria: 'Caladea',
      'Courier New': 'Liberation Mono',
    });

    expect(getSubstituteFontFamily('Times New Roman')).toBe('Liberation Serif');
    expect(getSubstituteFontFamily('times new roman')).toBe('Liberation Serif');
    expect(getSubstituteFontFamily('Arial')).toBe('Liberation Sans');
    expect(getSubstituteFontFamily('Calibri')).toBe('Carlito');
    expect(getSubstituteFontFamily('Cambria')).toBe('Caladea');
    expect(getSubstituteFontFamily('Courier New')).toBe('Liberation Mono');

    expect(resolveRenderFontStack('Times New Roman')).toBe(
      '"Liberation Serif", "Times New Roman", serif',
    );
    expect(resolveRenderFontStack('Arial')).toBe('"Liberation Sans", "Arial", sans-serif');
    expect(resolveRenderFontStack('Calibri')).toBe('"Carlito", "Calibri", sans-serif');
    expect(resolveRenderFontStack('Cambria')).toBe('"Caladea", "Cambria", serif');
    expect(resolveRenderFontStack('Courier New')).toBe(
      '"Liberation Mono", "Courier New", monospace',
    );
  });

  it('recovers original canonical Word font names from render stacks and blocks app shell UI fonts', () => {
    expect(canonicalizeWordFontFamily('"Liberation Serif", "Times New Roman", serif')).toBe(
      'Times New Roman',
    );
    expect(canonicalizeWordFontFamily('"Carlito", "Calibri", sans-serif')).toBe('Calibri');
    expect(canonicalizeWordFontFamily('"Liberation Sans", "Arial", sans-serif')).toBe('Arial');
    expect(canonicalizeWordFontFamily('"Caladea", "Cambria", serif')).toBe('Cambria');
    expect(canonicalizeWordFontFamily('"Liberation Mono", "Courier New", monospace')).toBe(
      'Courier New',
    );

    expect(isAppShellUiFont('Plus Jakarta Sans')).toBe(true);
    expect(isAppShellUiFont('DM Sans')).toBe(true);
    expect(canonicalizeWordFontFamily('Plus Jakarta Sans')).toBeNull();
    expect(canonicalizeWordFontFamily('"DM Sans", sans-serif')).toBeNull();

    // Never leak Plus Jakarta Sans or DM Sans into document render stacks
    const fallbackStack = resolveRenderFontStack('"Plus Jakarta Sans", "DM Sans", sans-serif');
    expect(fallbackStack).not.toMatch(/Plus Jakarta|DM Sans/i);
  });

  it('rewrites docx-preview stylesheet font-family declarations using the substitution map', () => {
    const sampleCss = `
      .esz-docx-wrapper { font-family: Calibri; font-size: 11pt; }
      .esz-docx span.t { font-family: "Times New Roman", serif !important; }
      .esz-docx code { font-family: "Courier New"; }
      .esz-docx h1 { font-family: inherit; }
    `;
    const rewritten = rewriteCssFontFamiliesForRender(sampleCss);
    expect(rewritten).toContain('font-family: "Carlito", "Calibri", sans-serif;');
    expect(rewritten).toContain('font-family: "Liberation Serif", "Times New Roman", serif !important;');
    expect(rewritten).toContain('font-family: "Liberation Mono", "Courier New", monospace;');
    expect(rewritten).toContain('font-family: inherit;');
  });

  it('never applies Plus Jakarta Sans or DM Sans inside DOC_THEMES or document font presets', () => {
    for (const preset of DOCUMENT_FONT_PRESETS) {
      expect(preset).not.toMatch(/Plus Jakarta|DM Sans/i);
    }
    for (const theme of Object.values(DOC_THEMES)) {
      expect(theme.fontFamily).not.toMatch(/Plus Jakarta|DM Sans/i);
      expect(theme.canonicalFontFamily).not.toMatch(/Plus Jakarta|DM Sans/i);
    }
  });

  it('preserves original Word font names in exported .docx (never writing render substitutes or app fonts)', async () => {
    const html = `
      <p><span data-doc-font="Times New Roman" style="font-family: &quot;Liberation Serif&quot;, &quot;Times New Roman&quot;, serif; font-size: 10pt">Academic text</span></p>
      <p><span data-doc-font="Calibri" style="font-family: &quot;Carlito&quot;, &quot;Calibri&quot;, sans-serif">Modern text</span></p>
      <p><span style="font-family: &quot;Liberation Sans&quot;, &quot;Arial&quot;, sans-serif">Arial run</span></p>
      <p><span style="font-family: &quot;Caladea&quot;, &quot;Cambria&quot;, serif">Cambria run</span></p>
      <p><span style="font-family: &quot;Liberation Mono&quot;, &quot;Courier New&quot;, monospace">Code run</span></p>
    `;

    const bytes = await htmlToDocxBytes(html, {
      ...DEFAULT_DOC_LAYOUT,
      docDefaultFont: 'Times New Roman',
      docDefaultFontSizePt: 10,
    });

    const zip = await JSZip.loadAsync(bytes);
    const docXml = (await zip.file('word/document.xml')?.async('string')) ?? '';
    const stylesXml = (await zip.file('word/styles.xml')?.async('string')) ?? '';
    const combinedXml = `${docXml}\n${stylesXml}`;

    expect(combinedXml).toContain('Times New Roman');
    expect(combinedXml).toContain('Calibri');
    expect(combinedXml).toContain('Arial');
    expect(combinedXml).toContain('Cambria');
    expect(combinedXml).toContain('Courier New');

    // Substituted render-only font names and app UI fonts must NEVER appear in exported .docx
    expect(combinedXml).not.toContain('Liberation Serif');
    expect(combinedXml).not.toContain('Liberation Sans');
    expect(combinedXml).not.toContain('Liberation Mono');
    expect(combinedXml).not.toContain('Carlito');
    expect(combinedXml).not.toContain('Caladea');
    expect(combinedXml).not.toMatch(/Plus Jakarta|DM Sans/i);
  });

  it('reads docDefaults / Normal typography from calibri-sample.docx and resolves theme fonts', async () => {
    const buf = fs.readFileSync(CALIBRI_FIXTURE_PATH);
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const defaults = await extractDocxDefaultTypography(ab);

    expect(defaults.fontFamily).toBe('Calibri');
    expect(defaults.fontSizePt).toBe(11);

    // Also verify theme font resolution (w:asciiTheme="minorHAnsi" -> word/theme/theme1.xml)
    const themeStylesXml = `
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults>
          <w:rPrDefault>
            <w:rPr>
              <w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>
              <w:sz w:val="20"/>
            </w:rPr>
          </w:rPrDefault>
        </w:docDefaults>
      </w:styles>
    `;
    const themeXml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements>
          <a:fontScheme name="Office">
            <a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
          </a:fontScheme>
        </a:themeElements>
      </a:theme>
    `;
    const resolved = parseDocDefaultsFromStyles(themeStylesXml, themeXml);
    expect(resolved.fontFamily).toBe('Calibri');
    expect(resolved.fontSizePt).toBe(10);
  });
});
