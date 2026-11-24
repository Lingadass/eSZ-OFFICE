import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import {describe, expect, it} from 'vitest';
import {importDocxToHtml} from './docxImport';
import {exportHtmlToDocxBytes} from './docxExport';
import {DEFAULT_DOC_LAYOUT, type DocPageLayout} from './docsTypes';

const FIXTURES_DIR = path.resolve(process.cwd(), 'tests/fixtures');
const CALIBRI_FIXTURE = path.join(FIXTURES_DIR, 'calibri-sample.docx');
const AIML_FIXTURE = path.join(FIXTURES_DIR, 'AIML.docx');
const hasAimlFixture = fs.existsSync(AIML_FIXTURE);

function readArrayBuffer(filePath: string): ArrayBuffer {
  const buf = fs.readFileSync(filePath);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

// 1x1 valid PNG bytes for synthetic tests
const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2S8AAAAASUVORK5CYII=';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
  const out = new ArrayBuffer(u8.byteLength);
  new Uint8Array(out).set(u8);
  return out;
}

async function buildSyntheticDocx(parts: Record<string, string | Uint8Array>): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  for (const [name, content] of Object.entries(parts)) {
    zip.file(name, content);
  }
  const u8 = await zip.generateAsync({type: 'uint8array'});
  return toArrayBuffer(u8);
}

describe('Step 3: High-fidelity OOXML Edit Import & Round-Trip Export', () => {
  it('imports calibri-sample.docx preserving fonts, headings, lists, shaded fixed-width table, image, and header/footer', async () => {
    const ab = readArrayBuffer(CALIBRI_FIXTURE);
    const res = await importDocxToHtml(ab);

    // Document defaults from styles.xml
    expect(res.layoutPatch.docDefaultFont).toBe('Calibri');
    expect(res.layoutPatch.docDefaultFontSizePt).toBe(11);
    expect(res.html).not.toContain('Plus Jakarta Sans');
    expect(res.html).not.toContain('DM Sans');

    // Headings & lists
    expect(res.html).toContain('<h1');
    expect(res.html).toContain('<h2');
    expect(res.html).toContain('<ul');
    expect(res.html).toContain('<ol');

    // Fixed-width table with column widths and header shading
    expect(res.stats.tables).toBeGreaterThanOrEqual(1);
    expect(res.html).toContain('<table');
    expect(res.html).toContain('colwidth=');
    expect(res.html.toUpperCase()).toContain('#DBEAFE');

    // Centred embedded PNG image with exact 240x120 dimensions
    expect(res.stats.images).toBeGreaterThanOrEqual(1);
    expect(res.html).toContain('data:image/png;base64,');
    expect(res.html).toContain('data-width="240"');
    expect(res.html).toContain('data-height="120"');
    expect(res.html).toContain('data-align="center"');

    // Header & footer page numbers
    expect(res.layoutPatch.headerText).toBeTruthy();
    expect(res.layoutPatch.showPageNumbers).toBe(true);

    // Clean document has zero locked/lossy elements
    expect(res.hasLockedElements).toBe(false);
    expect(res.hasLossyElements).toBe(false);
  });

  it('imports run/paragraph styles, merged tables, page borders, and Choice-only rounded rectangles without Fallback duplication', async () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:sz w:val="20"/>
        <w:color w:val="000000"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CustomLead">
    <w:name w:val="CustomLead"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="120" w:after="160" w:line="360" w:lineRule="auto"/>
      <w:keepNext/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="28"/>
      <w:color w:val="1E3A8A"/>
    </w:rPr>
  </w:style>
</w:styles>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="CustomLead"/>
        <w:ind w:left="360" w:firstLine="180"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:caps/>
          <w:spacing w:val="20"/>
          <w:highlight w:val="yellow"/>
          <w:u w:val="single"/>
        </w:rPr>
        <w:t>Laboratory Record</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
        <w:t>2</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:r>
        <mc:AlternateContent>
          <mc:Choice Requires="wps">
            <w:drawing>
              <wp:anchor>
                <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
                <wp:extent cx="4572000" cy="1143000"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp>
                      <wps:spPr>
                        <a:prstGeom prst="roundRect"/>
                        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                        <a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                      </wps:spPr>
                      <wps:txbx>
                        <w:txbxContent>
                          <w:p><w:r><w:t>NAME : A. STUDENT</w:t></w:r></w:p>
                          <w:p><w:r><w:t>REGISTER NO : 2025001</w:t></w:r></w:p>
                          <w:p><w:r><w:t>YEAR : III</w:t></w:r></w:p>
                        </w:txbxContent>
                      </wps:txbx>
                    </wps:wsp>
                  </a:graphicData>
                </a:graphic>
              </wp:anchor>
            </w:drawing>
          </mc:Choice>
          <mc:Fallback>
            <w:pict>
              <v:roundrect>
                <v:textbox>
                  <w:txbxContent>
                    <w:p><w:r><w:t>NAME : A. STUDENT</w:t></w:r></w:p>
                  </w:txbxContent>
                </v:textbox>
              </v:roundrect>
            </w:pict>
          </mc:Fallback>
        </mc:AlternateContent>
      </w:r>
    </w:p>

    <w:tbl>
      <w:tblPr>
        <w:jc w:val="center"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="6" w:color="000000"/>
          <w:bottom w:val="single" w:sz="6" w:color="000000"/>
          <w:left w:val="single" w:sz="6" w:color="000000"/>
          <w:right w:val="single" w:sz="6" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:color="000000"/>
          <w:insideV w:val="single" w:sz="4" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="2160"/>
        <w:gridCol w:w="2880"/>
        <w:gridCol w:w="2160"/>
      </w:tblGrid>
      <w:tr>
        <w:trPr><w:tblHeader/></w:trPr>
        <w:tc>
          <w:tcPr>
            <w:gridSpan w:val="2"/>
            <w:shd w:val="clear" w:fill="E2E8F0"/>
            <w:vAlign w:val="center"/>
          </w:tcPr>
          <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Merged Header</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>Span Vertically</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
      </w:tr>
    </w:tbl>

    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>
      <w:pgBorders w:offsetFrom="page">
        <w:top w:val="single" w:sz="12" w:space="24" w:color="000000"/>
        <w:bottom w:val="single" w:sz="12" w:space="24" w:color="000000"/>
        <w:left w:val="single" w:sz="12" w:space="24" w:color="000000"/>
        <w:right w:val="single" w:sz="12" w:space="24" w:color="000000"/>
      </w:pgBorders>
    </w:sectPr>
  </w:body>
</w:document>`;

    const ab = await buildSyntheticDocx({
      'word/styles.xml': stylesXml,
      'word/document.xml': documentXml,
    });
    const res = await importDocxToHtml(ab);

    // Default font & page borders
    expect(res.layoutPatch.docDefaultFont).toBe('Times New Roman');
    expect(res.layoutPatch.docDefaultFontSizePt).toBe(10);
    expect(res.layoutPatch.size).toBe('a4');
    expect(res.layoutPatch.pageBorderConfig?.offsetFrom).toBe('page');
    expect(res.layoutPatch.pageBorderConfig?.top?.color).toBe('#000000');

    // Inherited paragraph & run attributes
    expect(res.html).toContain('data-style-id="CustomLead"');
    expect(res.html).toContain('text-align: center');
    expect(res.html).toContain('data-keep-next="true"');
    expect(res.html).toContain('text-transform: uppercase');
    expect(res.html).toContain('letter-spacing: 1pt');
    expect(res.html).toMatch(/<sup>[\s\S]*2[\s\S]*<\/sup>/);
    expect(res.html).toContain('<mark');

    // mc:AlternateContent Choice-only rounded rectangle — ZERO duplication from mc:Fallback
    const nameMatches = res.html.match(/NAME : A\. STUDENT/g) ?? [];
    expect(nameMatches.length).toBe(1);
    expect(res.html).toContain('data-doc-textbox="true"');
    expect(res.html).toContain('data-box-radius="20"');
    expect(res.html).toContain('data-box-border-color="#000000"');
    expect(res.html).toContain('REGISTER NO : 2025001');
    expect(res.html).toContain('YEAR : III');
    expect(res.stats.alternateContentResolved).toBe(1);
    expect(res.stats.editableShapes).toBe(1);

    // Table merges (colspan=2, rowspan=2) & styling
    expect(res.html).toContain('colspan="2"');
    expect(res.html).toContain('rowspan="2"');
    expect(res.html).toContain('data-valign="middle"');
    expect(res.html.toUpperCase()).toContain('#E2E8F0');
    expect(res.html).toContain('data-table-align="center"');
  });

  it('logs warnings and creates locked fallback blocks for unsupported Word elements (SmartArt, Charts, VBA)', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <w:body>
    <w:p><w:r><w:t>Document with advanced elements</w:t></w:r></w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="3657600" cy="1828800"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="3657600" cy="1828800"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const ab = await buildSyntheticDocx({
      'word/document.xml': documentXml,
      'word/vbaProject.bin': new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]),
    });

    const res = await importDocxToHtml(ab);
    expect(res.hasLockedElements).toBe(true);
    expect(res.hasLossyElements).toBe(true);
    expect(res.html).toContain('data-doc-locked="true"');

    const codes = res.warnings.map((w) => w.code);
    expect(codes).toContain('smartart');
    expect(codes).toContain('chart');
    expect(codes).toContain('vba-macros');
  });

  it('round-trips imported .docx through exportHtmlToDocxBytes preserving canonical fonts, page border, table widths, image, and header/footer', async () => {
    // 1. Round-trip calibri-sample.docx
    const orig = await importDocxToHtml(readArrayBuffer(CALIBRI_FIXTURE));
    const layout1: DocPageLayout = {...DEFAULT_DOC_LAYOUT, ...orig.layoutPatch};
    const exportedBytes = await exportHtmlToDocxBytes(orig.html, layout1, []);

    // Inspect raw exported ZIP XML to verify canonical Word fonts only (no Carlito or app fonts)
    const exportedZip = await JSZip.loadAsync(exportedBytes);
    const stylesXml = (await exportedZip.file('word/styles.xml')?.async('string')) ?? '';
    const docXml = (await exportedZip.file('word/document.xml')?.async('string')) ?? '';
    expect(stylesXml).toContain('Calibri');
    expect(stylesXml).not.toContain('Carlito');
    expect(stylesXml).not.toContain('Plus Jakarta Sans');
    expect(docXml).not.toMatch(/w:ascii="Carlito"/i);
    expect(docXml).not.toMatch(/w:hAnsi="Carlito"/i);
    expect(docXml).toContain('<w:tblGrid>');
    expect(docXml.toUpperCase()).toContain('DBEAFE');

    // Re-import exported bytes
    const reimported = await importDocxToHtml(toArrayBuffer(exportedBytes));
    expect(reimported.layoutPatch.docDefaultFont).toBe('Calibri');
    expect(reimported.stats.tables).toBeGreaterThanOrEqual(1);
    expect(reimported.stats.images).toBeGreaterThanOrEqual(1);
    expect(reimported.layoutPatch.showPageNumbers).toBe(true);

    // 2. Round-trip Times New Roman + page border + rounded box + image
    const tnrHtml = `
      <p style="text-align: center"><span data-doc-font="Times New Roman" style="font-family: &quot;Liberation Serif&quot;, serif; font-size: 12pt"><strong>DEPARTMENT OF AI &amp; ML</strong></span></p>
      <p style="text-align: center"><img src="data:image/png;base64,${TINY_PNG_BASE64}" data-width="120" data-height="120" data-align="center" /></p>
      <div data-doc-textbox="true" data-box-custom="true" data-box-width="360" data-box-align="center" data-box-radius="12" data-box-border-color="#000000" data-box-border-width="2" data-box-border-style="solid" data-box-fill="#FFFFFF">
        <p><span data-doc-font="Times New Roman" style="font-size: 10pt">NAME : SAMPLE</span></p>
        <p><span data-doc-font="Times New Roman" style="font-size: 10pt">REGISTER NO : 12345</span></p>
      </div>
    `;
    const tnrLayout: DocPageLayout = {
      ...DEFAULT_DOC_LAYOUT,
      size: 'a4',
      docDefaultFont: 'Times New Roman',
      docDefaultFontSizePt: 10,
      headerText: 'AIML Record',
      footerText: 'Confidential',
      showPageNumbers: true,
      pageBorder: 'box',
      pageBorderConfig: {
        offsetFrom: 'page',
        display: 'allPages',
        top: {val: 'single', szPt: 1.5, color: '#000000', spacePt: 24},
        right: {val: 'single', szPt: 1.5, color: '#000000', spacePt: 24},
        bottom: {val: 'single', szPt: 1.5, color: '#000000', spacePt: 24},
        left: {val: 'single', szPt: 1.5, color: '#000000', spacePt: 24},
      },
    };
    const tnrBytes = await exportHtmlToDocxBytes(tnrHtml, tnrLayout, []);
    const tnrZip = await JSZip.loadAsync(tnrBytes);
    const tnrStyles = (await tnrZip.file('word/styles.xml')?.async('string')) ?? '';
    const tnrDocXml = (await tnrZip.file('word/document.xml')?.async('string')) ?? '';
    expect(tnrStyles).toContain('Times New Roman');
    expect(tnrStyles).not.toContain('Liberation Serif');
    expect(tnrDocXml).toContain('<w:pgBorders');

    const tnrReimport = await importDocxToHtml(toArrayBuffer(tnrBytes));
    expect(tnrReimport.layoutPatch.docDefaultFont).toBe('Times New Roman');
    expect(tnrReimport.layoutPatch.docDefaultFontSizePt).toBe(10);
    expect(tnrReimport.layoutPatch.pageBorderConfig?.top?.color).toBe('#000000');
    expect(tnrReimport.layoutPatch.headerText).toContain('AIML Record');
    expect(tnrReimport.html).toContain('DEPARTMENT OF AI &amp; ML');
    expect(tnrReimport.html).toContain('NAME : SAMPLE');
    expect(tnrReimport.stats.images).toBe(1);
  });

  describe.skipIf(!hasAimlFixture)('AIML.docx Edit-Mode Import Fixture', () => {
    it('imports AIML.docx preserving Times New Roman 10pt, black page border, 6-column table, logo, and non-duplicated rounded box', async () => {
      const ab = readArrayBuffer(AIML_FIXTURE);
      const res = await importDocxToHtml(ab);
      expect(res.layoutPatch.docDefaultFont).toBe('Times New Roman');
      expect(res.layoutPatch.docDefaultFontSizePt).toBe(10);
      expect(res.layoutPatch.pageBorderConfig).toBeTruthy();
      expect(res.stats.tables).toBeGreaterThanOrEqual(1);
      expect(res.stats.images).toBeGreaterThanOrEqual(1);
      expect(res.html).toContain('NAME');
      expect(res.html).toContain('REGISTER');
    });
  });

  describe('Step B: Edit mode fidelity (styles.xml, run inheritance, tabs, page breaks, floating images)', () => {
    it('uses styles.xml (Heading1, Heading3, BodyText, docDefaults) with black #000000 headings and never leaks app colors/fonts', async () => {
      const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:sz w:val="24"/>
        <w:color w:val="000000"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="000000"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="000000"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>`;

      const documentXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>ANNA UNIVERSITY</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="BodyText"/></w:pPr>
      <w:r><w:t>Bonafide internship report text.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading3"/></w:pPr>
      <w:r><w:t>SIGNATURE</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`;

      const ab = await buildSyntheticDocx({
        'word/styles.xml': stylesXml,
        'word/document.xml': documentXml,
      });
      const res = await importDocxToHtml(ab);

      expect(res.layoutPatch.hasImportedStyles).toBe(true);
      expect(res.layoutPatch.docDefaultFont).toBe('Times New Roman');
      expect(res.layoutPatch.docDefaultFontSizePt).toBe(12);
      expect(res.layoutPatch.docDefaultColor).toBe('#000000');

      // Heading1 ("ANNA UNIVERSITY") & Heading3 ("SIGNATURE") must keep explicit black #000000 & sizes from styles.xml
      expect(res.html).toMatch(/<h1[^>]*data-style-id="Heading1"[^>]*>[\s\S]*color: #000000[\s\S]*ANNA UNIVERSITY[\s\S]*<\/h1>/);
      expect(res.html).toMatch(/<h1[^>]*>[\s\S]*font-size: 16pt[\s\S]*ANNA UNIVERSITY[\s\S]*<\/h1>/);
      expect(res.html).toMatch(/<h3[^>]*data-style-id="Heading3"[^>]*>[\s\S]*color: #000000[\s\S]*SIGNATURE[\s\S]*<\/h3>/);
      expect(res.html).toMatch(/<h3[^>]*>[\s\S]*font-size: 12pt[\s\S]*SIGNATURE[\s\S]*<\/h3>/);
      expect(res.html).not.toContain('#1E3A8A');
      expect(res.html).not.toContain('Plus Jakarta Sans');
    });

    it('does not leak paragraph-mark w:pPr/w:rPr italic onto text runs and imports w:tab as real docTab nodes with tab stops', async () => {
      const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

      const settingsXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

      const documentXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:tabs>
          <w:tab w:val="right" w:pos="9026" w:leader="dot"/>
        </w:tabs>
        <w:rPr><w:i/></w:rPr>
      </w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>INTERNAL EXAMINER</w:t></w:r>
      <w:r><w:tab/></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>EXTERNAL EXAMINER</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

      const ab = await buildSyntheticDocx({
        'word/styles.xml': stylesXml,
        'word/settings.xml': settingsXml,
        'word/document.xml': documentXml,
      });
      const res = await importDocxToHtml(ab);

      // Must be bold only — NEVER italic leaked from w:pPr/w:rPr
      expect(res.html).toContain('<strong>INTERNAL EXAMINER</strong>');
      expect(res.html).toContain('<strong>EXTERNAL EXAMINER</strong>');
      expect(res.html).not.toContain('<em>');

      // Must emit a real docTab span with tab stops, NOT nbsp spaces
      expect(res.html).not.toContain('\u00A0\u00A0\u00A0\u00A0');
      expect(res.html).toContain('data-doc-tab="true"');
      expect(res.html).toContain('data-tab-align="right"');
      expect(res.html).toContain('data-tab-leader="dot"');
      expect(res.html).toContain('data-tab-pos="9026"');
      expect(res.html).toContain('data-tab-stops=');
      expect(res.layoutPatch.defaultTabStopPx).toBe(48);
    });

    it('imports w:br type="page" at its exact position before trailing text in the same paragraph', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Page 1 Footer Line</w:t></w:r></w:p>
    <w:p>
      <w:r>
        <w:br w:type="page"/>
        <w:t>BONAFIDE CERTIFICATE</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

      const ab = await buildSyntheticDocx({'word/document.xml': documentXml});
      const res = await importDocxToHtml(ab);

      const breakIdx = res.html.indexOf('data-page-break="true"');
      const certIdx = res.html.indexOf('BONAFIDE CERTIFICATE');
      expect(breakIdx).toBeGreaterThan(-1);
      expect(certIdx).toBeGreaterThan(breakIdx);
    });

    it('preserves centred anchored logos with vertical spacing and off-centre floating images with wrap/position attributes', async () => {
      const pngBytes = Uint8Array.from(Buffer.from(TINY_PNG_BASE64, 'base64'));
      const relsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>
  <Relationship Id="rIdSide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/side.png"/>
</Relationships>`;

      // Page width 11906 twips = 794px, margins 1440 twips = 96px -> content width 602px.
      // Logo width 160px = 1524000 EMU. Centred column offset = (602 - 160)/2 = 221px = 2105025 EMU.
      // Off-centre right square-wrapped image at column offset 450px = 4286250 EMU.
      const documentXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:t>ELECTRONICS AND COMMUNICATION ENGINEERING</w:t></w:r>
      <w:r>
        <w:drawing>
          <wp:anchor distT="114300" distB="114300">
            <wp:positionH relativeFrom="column"><wp:posOffset>2105025</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>228600</wp:posOffset></wp:positionV>
            <wp:extent cx="1524000" cy="1524000"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="1" name="College Logo"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic><pic:blipFill><a:blip r:embed="rIdLogo"/></pic:blipFill></pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:anchor distT="0" distB="95250">
            <wp:positionH relativeFrom="column"><wp:posOffset>4286250</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="1143000" cy="1143000"/>
            <wp:wrapSquare wrapText="bothSides"/>
            <wp:docPr id="2" name="Side Stamp"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic><pic:blipFill><a:blip r:embed="rIdSide"/></pic:blipFill></pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
      <w:r><w:t>Paragraph next to stamp</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`;

      const ab = await buildSyntheticDocx({
        'word/document.xml': documentXml,
        'word/_rels/document.xml.rels': relsXml,
        'word/media/logo.png': pngBytes,
        'word/media/side.png': pngBytes,
      });
      const res = await importDocxToHtml(ab);

      // Centred logo appears after heading text, centred, with vertical spacing
      const headingIdx = res.html.indexOf('ELECTRONICS AND COMMUNICATION ENGINEERING');
      const logoIdx = res.html.indexOf('alt="College Logo"');
      expect(headingIdx).toBeGreaterThan(-1);
      expect(logoIdx).toBeGreaterThan(headingIdx);
      expect(res.html).toMatch(/alt="College Logo"[^>]*data-align="center"/);
      expect(res.html).toMatch(/alt="College Logo"[^>]*data-margin-top="24"/);
      expect(res.html).toMatch(/alt="College Logo"[^>]*data-margin-bottom="12"/);

      // Off-centre stamp keeps floating + wrapSquare + float-left/right metadata
      expect(res.html).toMatch(/alt="Side Stamp"[^>]*data-floating="true"/);
      expect(res.html).toMatch(/alt="Side Stamp"[^>]*data-wrap="square"/);
      expect(res.html).toMatch(/alt="Side Stamp"[^>]*data-align="right"/);
      expect(res.html).toMatch(/alt="Side Stamp"[^>]*data-float-left="450"/);
    });
  });
});
