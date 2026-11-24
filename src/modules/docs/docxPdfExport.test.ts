import assert from 'node:assert/strict';
import {execFileSync} from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import {PDFDict, PDFDocument, PDFName, PDFRef} from 'pdf-lib';
import {describe, it} from 'vitest';
import {htmlToPdfBytes} from './docsPure';
import {DEFAULT_DOC_LAYOUT, type DocPageLayout} from './docsTypes';
import {importDocxToHtml} from './docxImport';
import {buildEmbeddedFontFaceCss, resolveBundledPdfFontFamily} from './pdfFonts';

const FIXTURES_DIR = path.resolve(process.cwd(), 'tests/fixtures');
const VELAVAN_ORIG_FIXTURE = path.join(FIXTURES_DIR, 'VELAVAN_FRONT_PAGE_FOR_INTERSHIP.docx');
const VELAVAN_RESAVED_FIXTURE = path.join(FIXTURES_DIR, 'VELAVAN_FRONT_PAGE_FOR.docx');
const CALIBRI_FIXTURE = path.join(FIXTURES_DIR, 'calibri-sample.docx');

// Valid 4x4 PNG (red/blue square)
const SAMPLE_LOGO_PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xdd, 0x8d, 0xb0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82,
]);

function hasCommand(cmd: string): boolean {
  try {
    execFileSync('which', [cmd], {stdio: 'ignore'});
    return true;
  } catch {
    return false;
  }
}

interface PdfInspectionSummary {
  pageCount: number;
  baseFonts: string[];
  embeddedFontCount: number;
  unembeddedFontNames: string[];
  imageCountTotal: number;
  imagesPerPage: number[];
}

/**
 * Inspects PDF internal font dictionaries, FontFile2/FontFile3 embedding streams,
 * and per-page Image XObjects using pdf-lib.
 */
async function inspectPdfStructure(pdfBytes: Uint8Array): Promise<PdfInspectionSummary> {
  const doc = await PDFDocument.load(pdfBytes);
  const ctx = doc.context;
  const baseFonts: string[] = [];
  const unembeddedFontNames: string[] = [];
  let embeddedFontCount = 0;

  for (const [, obj] of ctx.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFDict)) continue;
    const typeVal = obj.get(PDFName.of('Type'));
    if (typeVal !== PDFName.of('Font')) continue;

    const baseFontObj = obj.get(PDFName.of('BaseFont'));
    const baseFontName = baseFontObj ? baseFontObj.toString().replace(/^\//, '') : 'Unknown';
    baseFonts.push(baseFontName);

    const subtype = obj.get(PDFName.of('Subtype'));
    // Type0 composite font wraps DescendantFonts -> CIDFont -> FontDescriptor
    if (subtype === PDFName.of('Type0')) {
      const descendants = ctx.lookup(obj.get(PDFName.of('DescendantFonts')));
      let hasEmbeddedStream = false;
      if (descendants && 'asArray' in descendants) {
        for (const item of (descendants as {asArray: () => unknown[]}).asArray()) {
          const cidFont = ctx.lookup(item as PDFRef);
          if (cidFont instanceof PDFDict) {
            const fd = ctx.lookup(cidFont.get(PDFName.of('FontDescriptor')));
            if (fd instanceof PDFDict) {
              if (
                fd.get(PDFName.of('FontFile2')) ||
                fd.get(PDFName.of('FontFile3')) ||
                fd.get(PDFName.of('FontFile'))
              ) {
                hasEmbeddedStream = true;
              }
            }
          }
        }
      }
      if (hasEmbeddedStream) embeddedFontCount += 1;
      else unembeddedFontNames.push(baseFontName);
    } else if (subtype !== PDFName.of('CIDFontType2') && subtype !== PDFName.of('CIDFontType0')) {
      const fd = ctx.lookup(obj.get(PDFName.of('FontDescriptor')));
      const hasEmbeddedStream =
        fd instanceof PDFDict &&
        Boolean(
          fd.get(PDFName.of('FontFile2')) ||
            fd.get(PDFName.of('FontFile3')) ||
            fd.get(PDFName.of('FontFile')),
        );
      if (hasEmbeddedStream) embeddedFontCount += 1;
      else unembeddedFontNames.push(baseFontName);
    }
  }

  const imagesPerPage: number[] = [];
  let imageCountTotal = 0;
  for (const page of doc.getPages()) {
    let countOnPage = 0;
    const resources = page.node.Resources();
    if (resources) {
      const xObjects = ctx.lookup(resources.get(PDFName.of('XObject')));
      if (xObjects instanceof PDFDict) {
        for (const [, val] of xObjects.entries()) {
          const xObj = ctx.lookup(val);
          if (xObj && 'dict' in xObj) {
            const dict = (xObj as {dict: PDFDict}).dict;
            if (dict.get(PDFName.of('Subtype')) === PDFName.of('Image')) {
              countOnPage += 1;
            }
          }
        }
      }
    }
    imagesPerPage.push(countOnPage);
    imageCountTotal += countOnPage;
  }

  return {
    pageCount: doc.getPageCount(),
    baseFonts,
    embeddedFontCount,
    unembeddedFontNames,
    imageCountTotal,
    imagesPerPage,
  };
}

/**
 * Extracts per-page selectable text using pdfjs-dist.
 */
async function extractPdfTextPerPage(pdfBytes: Uint8Array): Promise<string[]> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(pdfBytes),
    useWorkerFetch: false,
    disableFontFace: true,
  });
  const pdf = await loadingTask.promise;
  const pageTexts: string[] = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const joined = content.items
      .map((item) => ('str' in item ? item.str : ''))
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();
    pageTexts.push(joined);
  }
  await loadingTask.destroy();
  return pageTexts;
}

async function buildVelavanStyleTwoPageDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>
</Relationships>`,
  );
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:color w:val="000000"/>
        <w:sz w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="000000"/><w:sz w:val="32"/></w:rPr>
  </w:style>
</w:styles>`,
  );
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>INTERNSHIP REPORT</w:t></w:r></w:p></w:hdr>`,
  );
  zip.file(
    'word/footer1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>ANNA UNIVERSITY CHENNAI</w:t></w:r></w:p></w:ftr>`,
  );
  zip.file('word/media/logo.png', SAMPLE_LOGO_PNG);
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="32"/></w:rPr><w:t>ANNA UNIVERSITY</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>ELECTRONICS AND COMMUNICATION ENGINEERING</w:t></w:r>
      <w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>2103786</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>320000</wp:posOffset></wp:positionV><wp:extent cx="1600200" cy="1600200"/><wp:wrapTopAndBottom/><wp:docPr id="1" name="AnnaLogo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1600200" cy="1600200"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="both"/></w:pPr>
      <w:r><w:t>Certified internship project submitted in partial fulfillment of the requirements for the award of the degree of Bachelor of Engineering with </w:t></w:r>
      <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>underlined honors</w:t></w:r>
      <w:r><w:t>, </w:t></w:r>
      <w:r><w:rPr><w:strike/></w:rPr><w:t>draft</w:t></w:r>
      <w:r><w:t>, H</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r>
      <w:r><w:t>O and 10</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r>
      <w:r><w:t> precision metrics across multiple lines to verify justified paragraph alignment.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:br w:type="page"/></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>BONAFIDE CERTIFICATE</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9026"/></w:tabs></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>INTERNAL EXAMINER</w:t></w:r>
      <w:r><w:tab/></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>EXTERNAL EXAMINER</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="6" w:color="000000"/>
          <w:bottom w:val="single" w:sz="6" w:color="000000"/>
          <w:left w:val="single" w:sz="6" w:color="000000"/>
          <w:right w:val="single" w:sz="6" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:color="666666"/>
          <w:insideV w:val="single" w:sz="4" w:color="666666"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:shd w:val="clear" w:fill="E2E8F0"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>REGISTER NO</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>310621106001</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId2"/>
      <w:footerReference w:type="default" r:id="rId3"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
      <w:pgBorders w:offsetFrom="page">
        <w:top w:val="double" w:sz="12" w:space="24" w:color="000000"/>
        <w:left w:val="double" w:sz="12" w:space="24" w:color="000000"/>
        <w:bottom w:val="double" w:sz="12" w:space="24" w:color="000000"/>
        <w:right w:val="double" w:sz="12" w:space="24" w:color="000000"/>
      </w:pgBorders>
    </w:sectPr>
  </w:body>
</w:document>`,
  );
  const bytes = await zip.generateAsync({type: 'uint8array'});
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe('Step D: Embedded-font PDF export (no StandardFonts / Helvetica)', () => {
  it('maps Word font families to bundled TrueType equivalents and never uses app UI fonts', () => {
    assert.equal(resolveBundledPdfFontFamily('Times New Roman'), 'liberation-serif');
    assert.equal(resolveBundledPdfFontFamily('"Liberation Serif", "Times New Roman", serif'), 'liberation-serif');
    assert.equal(resolveBundledPdfFontFamily('Arial'), 'liberation-sans');
    assert.equal(resolveBundledPdfFontFamily('Courier New'), 'liberation-mono');
    assert.equal(resolveBundledPdfFontFamily('Calibri'), 'carlito');
    assert.equal(resolveBundledPdfFontFamily('Cambria'), 'caladea');
    // App shell UI fonts must be ignored and fall back to document serif
    assert.equal(resolveBundledPdfFontFamily('Plus Jakarta Sans'), 'liberation-serif');
    assert.equal(resolveBundledPdfFontFamily('DM Sans'), 'liberation-serif');

    const printCss = buildEmbeddedFontFaceCss('https://example.local');
    assert.ok(printCss.includes('font-family: "Liberation Serif"'));
    assert.ok(printCss.includes('font-family: "Carlito"'));
    assert.ok(printCss.includes('font-family: "Caladea"'));
    assert.ok(printCss.includes('https://example.local/fonts/liberation-serif-400-normal.woff2'));
  });

  it('exports a 2-page Velavan-style .docx with embedded Liberation Serif (no Helvetica), logo image on page 1, and certificate starting on page 2', async () => {
    const docxBuffer = await buildVelavanStyleTwoPageDocx();
    const imported = await importDocxToHtml(docxBuffer);
    const layout: DocPageLayout = {
      ...DEFAULT_DOC_LAYOUT,
      ...imported.layoutPatch,
      watermarkEnabled: true,
      watermarkText: 'CONFIDENTIAL',
    };

    const pdfBytes = await htmlToPdfBytes(imported.html, 'VELAVAN_FRONT_PAGE', layout);
    const header = new TextDecoder('latin1').decode(pdfBytes.slice(0, 5));
    assert.equal(header, '%PDF-');

    const info = await inspectPdfStructure(pdfBytes);
    assert.equal(info.pageCount, 2, `Expected exactly 2 pages, got ${info.pageCount}`);
    assert.ok(
      info.baseFonts.some((f) => /LiberationSerif/i.test(f)),
      `Expected embedded LiberationSerif font, found: ${info.baseFonts.join(', ')}`,
    );
    assert.ok(
      !info.baseFonts.some((f) => /Helvetica/i.test(f)),
      `Helvetica must never appear in exported PDF, found: ${info.baseFonts.join(', ')}`,
    );
    assert.equal(
      info.unembeddedFontNames.length,
      0,
      `All fonts must be embedded; unembedded: ${info.unembeddedFontNames.join(', ')}`,
    );
    assert.ok(info.imagesPerPage[0] >= 1, 'Expected logo image XObject embedded on page 1');

    const pageTexts = await extractPdfTextPerPage(pdfBytes);
    assert.equal(pageTexts.length, 2);
    assert.ok(pageTexts[0].includes('ANNA UNIVERSITY'), `Page 1 missing ANNA UNIVERSITY: ${pageTexts[0]}`);
    assert.ok(
      pageTexts[0].includes('ELECTRONICS AND COMMUNICATION ENGINEERING'),
      `Page 1 missing department title: ${pageTexts[0]}`,
    );
    assert.ok(
      !pageTexts[0].includes('BONAFIDE CERTIFICATE'),
      'BONAFIDE CERTIFICATE must NOT appear on page 1',
    );
    assert.ok(
      pageTexts[1].includes('BONAFIDE CERTIFICATE'),
      `Page 2 must start with BONAFIDE CERTIFICATE: ${pageTexts[1]}`,
    );
    assert.ok(pageTexts[1].includes('INTERNAL EXAMINER'), 'Page 2 missing INTERNAL EXAMINER');
    assert.ok(pageTexts[1].includes('EXTERNAL EXAMINER'), 'Page 2 missing EXTERNAL EXAMINER');
    assert.ok(pageTexts[1].includes('310621106001'), 'Page 2 missing table cell text');

    // Run Poppler pdffonts / pdfimages / pdftotext if installed on host
    if (hasCommand('pdffonts') && hasCommand('pdfimages') && hasCommand('pdftotext')) {
      const tmpPdf = path.join(os.tmpdir(), `esz-stepd-${Date.now()}.pdf`);
      try {
        fs.writeFileSync(tmpPdf, pdfBytes);
        const fontsOut = execFileSync('pdffonts', [tmpPdf], {encoding: 'utf8'});
        assert.match(fontsOut, /LiberationSerif/i);
        assert.doesNotMatch(fontsOut, /Helvetica/i);

        const imagesOut = execFileSync('pdfimages', ['-list', tmpPdf], {encoding: 'utf8'});
        assert.match(imagesOut, /\b1\s+0\s+image\b/);

        const textOut = execFileSync('pdftotext', [tmpPdf, '-'], {encoding: 'utf8'});
        assert.ok(textOut.includes('BONAFIDE CERTIFICATE'));
      } finally {
        if (fs.existsSync(tmpPdf)) fs.unlinkSync(tmpPdf);
      }
    }
  });

  it('exports tests/fixtures/calibri-sample.docx with embedded Carlito (no Helvetica) and extractable text', async () => {
    const raw = fs.readFileSync(CALIBRI_FIXTURE);
    const buf = raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength);
    const imported = await importDocxToHtml(buf);
    const layout: DocPageLayout = {
      ...DEFAULT_DOC_LAYOUT,
      ...imported.layoutPatch,
    };

    const pdfBytes = await htmlToPdfBytes(imported.html, 'calibri-sample', layout);
    const info = await inspectPdfStructure(pdfBytes);

    assert.ok(info.pageCount >= 1);
    assert.ok(
      info.baseFonts.some((f) => /Carlito/i.test(f)),
      `Expected embedded Carlito font for Calibri document, found: ${info.baseFonts.join(', ')}`,
    );
    assert.ok(
      !info.baseFonts.some((f) => /Helvetica/i.test(f)),
      `Helvetica must never appear in exported PDF, found: ${info.baseFonts.join(', ')}`,
    );
    assert.equal(info.unembeddedFontNames.length, 0);

    const pageTexts = await extractPdfTextPerPage(pdfBytes);
    assert.ok(pageTexts.join(' ').trim().length > 10, 'Expected extractable selectable text');
  });

  const hasVelavanFixtures =
    fs.existsSync(VELAVAN_ORIG_FIXTURE) && fs.existsSync(VELAVAN_RESAVED_FIXTURE);

  it.skipIf(!hasVelavanFixtures)(
    'exports both user-supplied Velavan fixtures with embedded Liberation Serif, 2 pages, logo image, and certificate on page 2',
    async () => {
      for (const fixturePath of [VELAVAN_ORIG_FIXTURE, VELAVAN_RESAVED_FIXTURE]) {
        const raw = fs.readFileSync(fixturePath);
        const buf = raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength);
        const imported = await importDocxToHtml(buf);
        const layout: DocPageLayout = {
          ...DEFAULT_DOC_LAYOUT,
          ...imported.layoutPatch,
        };
        const pdfBytes = await htmlToPdfBytes(
          imported.html,
          path.basename(fixturePath, '.docx'),
          layout,
        );

        const info = await inspectPdfStructure(pdfBytes);
        assert.equal(info.pageCount, 2, `${path.basename(fixturePath)}: expected 2 pages`);
        assert.ok(
          info.baseFonts.some((f) => /LiberationSerif/i.test(f)),
          `${path.basename(fixturePath)}: expected LiberationSerif, got ${info.baseFonts.join(', ')}`,
        );
        assert.ok(
          !info.baseFonts.some((f) => /Helvetica/i.test(f)),
          `${path.basename(fixturePath)}: Helvetica must not be used`,
        );
        assert.equal(info.unembeddedFontNames.length, 0);
        assert.ok(info.imageCountTotal >= 1, `${path.basename(fixturePath)}: expected logo image`);

        const pageTexts = await extractPdfTextPerPage(pdfBytes);
        assert.ok(
          /CERTIFICATE/i.test(pageTexts[1] ?? ''),
          `${path.basename(fixturePath)}: expected certificate on page 2`,
        );

        if (hasCommand('pdffonts') && hasCommand('pdfimages') && hasCommand('pdftotext')) {
          const tmpPdf = path.join(os.tmpdir(), `esz-fixture-${Date.now()}.pdf`);
          try {
            fs.writeFileSync(tmpPdf, pdfBytes);
            const fontsOut = execFileSync('pdffonts', [tmpPdf], {encoding: 'utf8'});
            assert.match(fontsOut, /LiberationSerif/i);
            assert.doesNotMatch(fontsOut, /Helvetica/i);

            const imagesOut = execFileSync('pdfimages', ['-list', tmpPdf], {encoding: 'utf8'});
            assert.match(imagesOut, /\bimage\b/i);

            const textOut = execFileSync('pdftotext', [tmpPdf, '-'], {encoding: 'utf8'});
            assert.ok(textOut.trim().length > 20);
          } finally {
            if (fs.existsSync(tmpPdf)) fs.unlinkSync(tmpPdf);
          }
        }
      }
    },
  );
});
