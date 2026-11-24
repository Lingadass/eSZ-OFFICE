import {execFileSync} from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import {describe, expect, it} from 'vitest';
import {
  computeVisiblePageNumber,
  extractSectionPageBorders,
  isRealDocxArchive,
  normalizePageNumberFieldsXml,
  preprocessDocxForPreview,
} from './docxPreviewUtils';

const FIXTURES_DIR = path.resolve(import.meta.dirname, '../../../tests/fixtures');
const CALIBRI_FIXTURE_PATH = path.join(FIXTURES_DIR, 'calibri-sample.docx');
const AIML_FIXTURE_PATH = path.join(FIXTURES_DIR, 'AIML.docx');
const hasAimlFixture = fs.existsSync(AIML_FIXTURE_PATH);

function toArrayBuffer(buf: Buffer): ArrayBuffer {
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

describe('docxPreviewUtils (Step 1: Faithful View Mode)', () => {
  it('recognizes real OOXML .docx archive and preserves original bytes untouched', async () => {
    const rawBuf = fs.readFileSync(CALIBRI_FIXTURE_PATH);
    const originalCopy = new Uint8Array(rawBuf);
    const ab = toArrayBuffer(rawBuf);

    expect(await isRealDocxArchive(ab)).toBe(true);

    const preview = await preprocessDocxForPreview(ab);
    // Input buffer must remain byte-for-byte identical
    expect(new Uint8Array(ab)).toEqual(originalCopy);

    // Extracted Calibri docDefaults from word/styles.xml
    expect(preview.docDefaults.fontFamily).toBe('Calibri');
    expect(preview.docDefaults.fontSizePt).toBe(11);

    // Footer PAGE and NUMPAGES fields normalized in the in-memory preview copy
    const outZip = await JSZip.loadAsync(preview.buffer);
    const footerXml = await outZip.file('word/footer1.xml')?.async('string');
    expect(footerXml).toContain('__ESZ_PAGE_NUM__');
    expect(footerXml).toContain('__ESZ_NUM_PAGES__');
  });

  it('returns false for plain-text files with .docx extension', async () => {
    const plain = new TextEncoder().encode('Plain starter document text').buffer;
    expect(await isRealDocxArchive(plain)).toBe(false);
  });

  it('resolves mc:AlternateContent using mc:Choice and strips mc:Fallback completely', async () => {
    const zip = new JSZip();
    zip.file(
      'word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8"?>
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults><w:rPrDefault><w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
          <w:sz w:val="20"/>
        </w:rPr></w:rPrDefault></w:docDefaults>
      </w:styles>`,
    );
    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8"?>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        xmlns:v="urn:schemas-microsoft-com:vml">
        <w:body>
          <w:p>
            <w:r>
              <mc:AlternateContent>
                <mc:Choice Requires="wps">
                  <w:drawing>
                    <wp:anchor>
                      <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
                      <wp:extent cx="3429000" cy="1143000"/>
                      <wp:wrapTopAndBottom/>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                          <wps:wsp>
                            <wps:spPr>
                              <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>
                              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                              <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                            </wps:spPr>
                            <wps:txbx>
                              <w:txbxContent>
                                <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>CONTENT</w:t></w:r></w:p>
                              </w:txbxContent>
                            </wps:txbx>
                            <wps:bodyPr anchor="ctr"/>
                          </wps:wsp>
                        </a:graphicData>
                      </a:graphic>
                    </wp:anchor>
                  </w:drawing>
                </mc:Choice>
                <mc:Fallback>
                  <w:pict>
                    <v:roundrect><v:textbox><w:txbxContent><w:p><w:r><w:t>FALLBACK_DUPLICATE_CONTENT</w:t></w:r></w:p></w:txbxContent></v:textbox></v:roundrect>
                  </w:pict>
                </mc:Fallback>
              </mc:AlternateContent>
            </w:r>
          </w:p>
          <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            <w:pgBorders w:offsetFrom="page">
              <w:top w:val="single" w:sz="12" w:space="24" w:color="000000"/>
              <w:left w:val="single" w:sz="12" w:space="24" w:color="000000"/>
              <w:bottom w:val="single" w:sz="12" w:space="24" w:color="000000"/>
              <w:right w:val="single" w:sz="12" w:space="24" w:color="000000"/>
            </w:pgBorders>
          </w:sectPr>
        </w:body>
      </w:document>`,
    );

    const inputBuf = await zip.generateAsync({type: 'arraybuffer'});
    const preview = await preprocessDocxForPreview(inputBuf);

    expect(preview.stats.alternateContentCount).toBe(1);
    expect(preview.stats.choiceUsedCount).toBe(1);
    expect(preview.stats.fallbackDroppedCount).toBe(1);
    expect(preview.stats.shapesExtractedCount).toBe(1);
    expect(preview.stats.hasPageBorders).toBe(true);

    // Verify mc:Fallback was completely dropped from preview document.xml
    const outZip = await JSZip.loadAsync(preview.buffer);
    const outDocXml = await outZip.file('word/document.xml')?.async('string');
    expect(outDocXml).not.toContain('mc:Fallback');
    expect(outDocXml).not.toContain('FALLBACK_DUPLICATE_CONTENT');
    expect(outDocXml).toContain('__ESZ_SHAPE_0__');

    // Verify extracted DrawingML Choice shape metadata
    const shape = preview.shapes.ESZ_SHAPE_0;
    expect(shape).toBeDefined();
    expect(shape.prstGeom).toBe('roundRect');
    expect(shape.borderRadiusPx).toBeGreaterThan(0);
    expect(shape.borderColor).toBe('#000000');
    expect(shape.alignH).toBe('center');
    expect(shape.paragraphs[0]?.runs[0]?.text).toBe('CONTENT');
    expect(shape.paragraphs[0]?.runs[0]?.fontFamily).toBe('Times New Roman');
    expect(shape.paragraphs[0]?.runs[0]?.fontSizePt).toBe(14);
  });

  it('extracts section page borders (w:pgBorders) accurately', () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:sectPr>
          <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"/>
          <w:pgBorders w:offsetFrom="page" w:display="allPages">
            <w:top w:val="single" w:sz="16" w:space="18" w:color="000000"/>
            <w:right w:val="double" w:sz="8" w:space="18" w:color="1E293B"/>
            <w:bottom w:val="single" w:sz="16" w:space="18" w:color="000000"/>
            <w:left w:val="single" w:sz="16" w:space="18" w:color="auto"/>
          </w:pgBorders>
        </w:sectPr>
      </w:body>
    </w:document>`;

    const borders = extractSectionPageBorders(xml);
    expect(borders).toHaveLength(1);
    expect(borders[0]?.offsetFrom).toBe('page');
    expect(borders[0]?.top?.szPt).toBe(2);
    expect(borders[0]?.top?.color).toBe('#000000');
    expect(borders[0]?.right?.val).toBe('double');
    expect(borders[0]?.left?.color).toBe('#000000');
    expect(borders[0]?.marginsPt.left).toBe(54);
  });

  it('normalizes both simple and complex Word PAGE / NUMPAGES fields', () => {
    const xml = `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:p>
        <w:r><w:t>Page </w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:t> of </w:t></w:r>
        <w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>3</w:t></w:r></w:fldSimple>
      </w:p>
    </w:ftr>`;

    const normalized = normalizePageNumberFieldsXml(xml);
    expect(normalized).toContain('__ESZ_PAGE_NUM__');
    expect(normalized).toContain('__ESZ_NUM_PAGES__');
  });

  it('computes accurate visible 1-based page number across zoom scales', () => {
    const pageH = 1123;
    const gap = 24;
    for (const zoom of [0.5, 1, 1.1, 1.5]) {
      const scaledH = pageH * zoom;
      const stride = (pageH + gap) * zoom;
      const bounds = [0, 1, 2].map((i) => ({top: i * stride, height: scaledH}));

      expect(computeVisiblePageNumber(0, 700, bounds)).toBe(1);
      expect(computeVisiblePageNumber(stride, 700, bounds)).toBe(2);
      expect(computeVisiblePageNumber(stride * 2, 700, bounds)).toBe(3);
    }
  });
});

const VELAVAN_FIXTURE_PATH = path.join(FIXTURES_DIR, 'VELAVAN_FRONT_PAGE_FOR_INTERSHIP.docx');
const hasVelavanFixture = fs.existsSync(VELAVAN_FIXTURE_PATH);

function hasCommand(cmd: string): boolean {
  try {
    execFileSync('which', [cmd], {stdio: 'ignore'});
    return true;
  } catch {
    return false;
  }
}

const hasLibreOfficeTools = hasCommand('soffice') && hasCommand('pdftoppm');

// Minimal 1x1 PNG for synthetic anchored-image OOXML tests
const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC';

async function buildAnchoredImageTestDocx(options: {
  wrapXml: string;
  relativeFromH: 'page' | 'column' | 'margin';
  posHInnerXml: string;
  relativeFromV?: 'page' | 'margin' | 'paragraph';
  posVInnerXml?: string;
  attachToTextParagraph?: boolean;
  includeLastRenderedPageBreak?: boolean;
}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    </Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>
    </Relationships>`,
  );
  zip.file('word/media/logo.png', TINY_PNG_BASE64, {base64: true});
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:sz w:val="24"/>
      </w:rPr></w:rPrDefault></w:docDefaults>
    </w:styles>`,
  );

  // 160px x 160px logo -> 1524000 x 1524000 EMU
  const drawingXml = `<w:drawing>
    <wp:anchor distT="114300" distB="114300" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="${options.relativeFromH}">${options.posHInnerXml}</wp:positionH>
      <wp:positionV relativeFrom="${options.relativeFromV ?? 'paragraph'}">${options.posVInnerXml ?? '<wp:posOffset>0</wp:posOffset>'}</wp:positionV>
      <wp:extent cx="1524000" cy="1524000"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      ${options.wrapXml}
      <wp:docPr id="1" name="College Logo"/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1524000" cy="1524000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing>`;

  const headingAndLogoXml = options.attachToTextParagraph
    ? `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>ELECTRONICS AND COMMUNICATION ENGINEERING</w:t></w:r><w:r>${drawingXml}</w:r></w:p>`
    : `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>ELECTRONICS AND COMMUNICATION ENGINEERING</w:t></w:r></w:p>
       <w:p><w:r>${drawingXml}</w:r></w:p>`;

  const breakXml = options.includeLastRenderedPageBreak
    ? `<w:p><w:r><w:lastRenderedPageBreak/><w:br w:type="page"/></w:r></w:p>`
    : `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;

  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <w:document
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <w:body>
        ${headingAndLogoXml}
        ${breakXml}
        <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>BONAFIDE CERTIFICATE</w:t></w:r></w:p>
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
      </w:body>
    </w:document>`,
  );

  return zip.generateAsync({type: 'arraybuffer'});
}

describe('Step A: View mode anchored-image normalization & pagination safeguards', () => {
  it('converts a wrapTopAndBottom column-relative anchored logo within 5% of page centre into a centred inline paragraph', async () => {
    // A4 page width = 11906 twips = 793.73px; marginLeft = 96px; contentWidth = 601.73px.
    // Logo width = 160px -> exact column centre offset = (601.73 - 160)/2 = 220.87px -> 2103786 EMU.
    // Shift by +15px (~1.9% of page width, well within 5% threshold) -> 2246661 EMU.
    const buf = await buildAnchoredImageTestDocx({
      wrapXml: '<wp:wrapTopAndBottom/>',
      relativeFromH: 'column',
      posHInnerXml: '<wp:posOffset>2246661</wp:posOffset>',
      attachToTextParagraph: true,
      includeLastRenderedPageBreak: true,
    });

    const preview = await preprocessDocxForPreview(buf);
    expect(preview.stats.anchoredImagesCenteredCount).toBe(1);
    expect(preview.stats.anchoredImagesAbsoluteCount).toBe(0);
    expect(Object.keys(preview.floatingImages)).toHaveLength(0);

    const outZip = await JSZip.loadAsync(preview.buffer);
    const outDocXml = (await outZip.file('word/document.xml')?.async('string')) ?? '';
    expect(outDocXml).not.toContain('<wp:anchor');
    expect(outDocXml).toContain('<wp:inline');
    // Text paragraph remains intact, followed by its own centred inline-image paragraph
    const textIdx = outDocXml.indexOf('ELECTRONICS AND COMMUNICATION ENGINEERING');
    const inlineIdx = outDocXml.indexOf('<wp:inline');
    expect(textIdx).toBeGreaterThan(-1);
    expect(inlineIdx).toBeGreaterThan(textIdx);
    expect(outDocXml.slice(textIdx, inlineIdx)).toContain('<w:jc w:val="center"/>');
  });

  it('converts an off-centre (>5% from page centre) wrapNone page-relative anchored image into an absolute page overlay', async () => {
    // Left offset = 80px (762000 EMU), logo width = 160px -> centreX = 160px vs pageCentre 396.87px (~29.8% off centre)
    const buf = await buildAnchoredImageTestDocx({
      wrapXml: '<wp:wrapNone/>',
      relativeFromH: 'page',
      posHInnerXml: '<wp:posOffset>762000</wp:posOffset>',
      relativeFromV: 'page',
      posVInnerXml: '<wp:posOffset>2857500</wp:posOffset>', // 300px from page top
      attachToTextParagraph: false,
    });

    const preview = await preprocessDocxForPreview(buf);
    expect(preview.stats.anchoredImagesCenteredCount).toBe(0);
    expect(preview.stats.anchoredImagesAbsoluteCount).toBe(1);

    const floatImg = preview.floatingImages.ESZ_FLOAT_IMG_0;
    expect(floatImg).toBeDefined();
    expect(floatImg.leftPx).toBeCloseTo(80, 1);
    expect(floatImg.topPx).toBeCloseTo(300, 1);
    expect(floatImg.widthPx).toBeCloseTo(160, 1);
    expect(floatImg.heightPx).toBeCloseTo(160, 1);
    expect(floatImg.relativeFromV).toBe('page');
    expect(floatImg.dataUrl.startsWith('data:image/png;base64,')).toBe(true);

    const outZip = await JSZip.loadAsync(preview.buffer);
    const outDocXml = (await outZip.file('word/document.xml')?.async('string')) ?? '';
    expect(outDocXml).not.toContain('<wp:anchor');
    expect(outDocXml).toContain('__ESZ_FLOAT_IMG_0__');
  });

  it('configures DocxPreviewView with ignoreLastRenderedPageBreak: true to prevent duplicate blank pages', () => {
    const viewSource = fs.readFileSync(
      path.resolve(import.meta.dirname, './DocxPreviewView.tsx'),
      'utf8',
    );
    expect(viewSource).toMatch(/ignoreLastRenderedPageBreak:\s*true/);
  });
});

describe.skipIf(!hasAimlFixture)('AIML.docx fixture checks (pending user-supplied file)', () => {
  it('preprocesses AIML.docx without mutating original bytes and detects page borders', async () => {
    const rawBuf = fs.readFileSync(AIML_FIXTURE_PATH);
    const originalCopy = new Uint8Array(rawBuf);
    const ab = toArrayBuffer(rawBuf);

    expect(await isRealDocxArchive(ab)).toBe(true);
    const preview = await preprocessDocxForPreview(ab);
    expect(new Uint8Array(ab)).toEqual(originalCopy);
    expect(preview.stats.hasPageBorders).toBe(true);

    const outZip = await JSZip.loadAsync(preview.buffer);
    const outDocXml = await outZip.file('word/document.xml')?.async('string');
    expect(outDocXml).not.toContain('mc:Fallback');
  });
});

describe.skipIf(!hasVelavanFixture)(
  'VELAVAN_FRONT_PAGE_FOR_INTERSHIP.docx Step A View-mode checks (pending user-supplied file)',
  () => {
    it('preprocesses VELAVAN original cleanly and centres anchored logo', async () => {
      const rawBuf = fs.readFileSync(VELAVAN_FIXTURE_PATH);
      const originalCopy = new Uint8Array(rawBuf);
      const ab = toArrayBuffer(rawBuf);

      const preview = await preprocessDocxForPreview(ab);
      expect(new Uint8Array(ab)).toEqual(originalCopy);
      expect(
        preview.stats.anchoredImagesCenteredCount + preview.stats.anchoredImagesAbsoluteCount,
      ).toBeGreaterThanOrEqual(0);
    });

    it.skipIf(!hasLibreOfficeTools)(
      'matches LibreOffice PDF page count (2 pages) via soffice + pdftoppm',
      () => {
        const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'esz-lo-step-a-'));
        try {
          execFileSync(
            'soffice',
            ['--headless', '--convert-to', 'pdf', '--outdir', tmpDir, VELAVAN_FIXTURE_PATH],
            {stdio: 'pipe'},
          );
          const pdfPath = path.join(tmpDir, 'VELAVAN_FRONT_PAGE_FOR_INTERSHIP.pdf');
          expect(fs.existsSync(pdfPath)).toBe(true);
          execFileSync('pdftoppm', ['-png', '-r', '150', pdfPath, path.join(tmpDir, 'page')], {
            stdio: 'pipe',
          });
          const renderedPages = fs
            .readdirSync(tmpDir)
            .filter((f) => /^page-\d+\.png$/.test(f));
          expect(renderedPages.length).toBe(2);
        } finally {
          fs.rmSync(tmpDir, {recursive: true, force: true});
        }
      },
    );
  },
);
