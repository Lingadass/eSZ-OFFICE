import {execFileSync} from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import {describe, expect, it} from 'vitest';
import {DEFAULT_DOC_LAYOUT, type DocPageLayout} from './docsTypes';
import {exportHtmlToDocxBytes} from './docxExport';
import {importDocxToHtml} from './docxImport';
import {patchOriginalDocxPackage} from './docxLosslessSave';

const FIXTURES_DIR = path.resolve(import.meta.dirname, '../../../tests/fixtures');
const CALIBRI_FIXTURE_PATH = path.join(FIXTURES_DIR, 'calibri-sample.docx');
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

const hasLibreOfficeTools = hasCommand('soffice') && hasCommand('pdftotext') && hasCommand('pdfimages');

const TINY_PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC';
// Distinct 1x1 blue PNG for testing newly inserted media
const BLUE_PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgYPgPAAEDAQAIicLsAAAAAElFTkSuQmCC';

function toArrayBuffer(buf: Buffer | Uint8Array): ArrayBuffer {
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

async function buildVelavanStylePackage(): Promise<{
  buffer: ArrayBuffer;
  parts: Record<string, string>;
}> {
  const zip = new JSZip();

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`;

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>
</Relationships>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:color w:val="000000"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="000000"/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="80" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="000000"/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>`;

  const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

  const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

  const fontTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Times New Roman"><w:family w:val="roman"/></w:font>
</w:fonts>`;

  const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="VelavanCustomTheme">
  <a:themeElements>
    <a:clrScheme name="Custom"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme>
    <a:fontScheme name="Custom"><a:majorFont><a:latin typeface="Times New Roman"/></a:majorFont><a:minorFont><a:latin typeface="Times New Roman"/></a:minorFont></a:fontScheme>
  </a:themeElements>
</a:theme>`;

  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>INTERNSHIP REPORT</w:t></w:r></w:p>
</w:hdr>`;

  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>ANNA UNIVERSITY CHENNAI</w:t></w:r></w:p>
</w:ftr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p w:rsidR="00AA1122"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>ANNA UNIVERSITY</w:t></w:r></w:p>
    <w:p w:rsidR="00BB3344"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>ELECTRONICS AND COMMUNICATION ENGINEERING</w:t></w:r><w:r><w:drawing><wp:anchor distT="114300" distB="114300" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>2103786</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV><wp:extent cx="1524000" cy="1524000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/><wp:docPr id="1" name="College Logo"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId8"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1524000" cy="1524000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>
    <w:p w:rsidR="00CC5566"><w:r><w:br w:type="page"/></w:r></w:p>
    <w:p w:rsidR="00DD7788"><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>BONAFIDE CERTIFICATE</w:t></w:r></w:p>
    <w:p w:rsidR="00EE9900"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>Certified that this internship report is the bonafide work of Velavan.</w:t></w:r></w:p>
    <w:p w:rsidR="00FF1122"><w:pPr><w:tabs><w:tab w:val="right" w:pos="9026"/></w:tabs><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>INTERNAL EXAMINER</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>EXTERNAL EXAMINER</w:t></w:r></w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rId6"/><w:footerReference w:type="default" r:id="rId7"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="12" w:space="24" w:color="000000"/><w:left w:val="single" w:sz="12" w:space="24" w:color="000000"/><w:bottom w:val="single" w:sz="12" w:space="24" w:color="000000"/><w:right w:val="single" w:sz="12" w:space="24" w:color="000000"/></w:pgBorders></w:sectPr>
  </w:body>
</w:document>`;

  const parts: Record<string, string> = {
    '[Content_Types].xml': contentTypesXml,
    '_rels/.rels': rootRelsXml,
    'word/_rels/document.xml.rels': docRelsXml,
    'word/styles.xml': stylesXml,
    'word/numbering.xml': numberingXml,
    'word/settings.xml': settingsXml,
    'word/fontTable.xml': fontTableXml,
    'word/theme/theme1.xml': themeXml,
    'word/header1.xml': headerXml,
    'word/footer1.xml': footerXml,
    'word/document.xml': documentXml,
  };

  for (const [p, content] of Object.entries(parts)) {
    zip.file(p, content);
  }
  zip.file('word/media/logo.png', TINY_PNG_B64, {base64: true});

  const buffer = await zip.generateAsync({type: 'arraybuffer'});
  return {buffer, parts};
}

describe('Step C — Lossless .docx save', () => {
  it('preserves all auxiliary package parts and writes every untouched paragraph, page break, tab stop, wp:anchor, and sectPr back verbatim', async () => {
    const {buffer: origBuf, parts} = await buildVelavanStylePackage();
    const origCopy = new Uint8Array(origBuf.slice(0));

    const imported = await importDocxToHtml(origBuf);
    const layout: DocPageLayout = {...DEFAULT_DOC_LAYOUT, ...imported.layoutPatch};

    // Simulate TipTap DOM normalization (rgb() color format, reordered nested tags, <br> vs <br />)
    const tiptapNormalizedHtml = imported.html
      .replace(/color:\s*#000000/gi, 'color: rgb(0, 0, 0)')
      .replace(/<br\s*\/>/g, '<br>');

    const result = await patchOriginalDocxPackage(tiptapNormalizedHtml, layout, [], origBuf);
    expect(result).not.toBeNull();
    const {bytes: savedBytes, stats} = result!;

    // Original caller ArrayBuffer must never be mutated
    expect(new Uint8Array(origBuf)).toEqual(origCopy);

    // Every single body block was preserved verbatim (0 regenerated blocks!)
    expect(stats.preservedBlocks).toBe(6);
    expect(stats.generatedBlocks).toBe(0);
    expect(stats.addedMediaCount).toBe(0);
    expect(stats.sectPrPreservedVerbatim).toBe(true);

    const outZip = await JSZip.loadAsync(savedBytes);

    // 1. Auxiliary package parts MUST be byte-for-byte identical
    for (const partPath of [
      '[Content_Types].xml',
      '_rels/.rels',
      'word/_rels/document.xml.rels',
      'word/styles.xml',
      'word/numbering.xml',
      'word/settings.xml',
      'word/fontTable.xml',
      'word/theme/theme1.xml',
      'word/header1.xml',
      'word/footer1.xml',
    ]) {
      const outContent = await outZip.file(partPath)?.async('string');
      expect(outContent).toBe(parts[partPath]);
    }

    // Original media untouched
    const outLogoB64 = await outZip.file('word/media/logo.png')?.async('base64');
    expect(outLogoB64).toBe(TINY_PNG_B64);

    // 2. document.xml preserves exact original paragraphs, rsidR, pStyle, wp:anchor, tabs, page break, and sectPr
    const outDocXml = (await outZip.file('word/document.xml')?.async('string')) ?? '';
    expect(outDocXml).toContain('w:rsidR="00AA1122"');
    expect(outDocXml).toContain('<w:pStyle w:val="Heading1"/>');
    expect(outDocXml).toContain('<w:pStyle w:val="Heading3"/>');
    expect(outDocXml).toContain('<w:pStyle w:val="BodyText"/>');
    expect(outDocXml).toContain('<w:br w:type="page"/>');
    expect(outDocXml).toContain('<w:tabs><w:tab w:val="right" w:pos="9026"/></w:tabs>');
    expect(outDocXml).toContain('<w:tab/>');
    expect(outDocXml).toContain('<wp:anchor');
    expect(outDocXml).toContain('<wp:posOffset>2103786</wp:posOffset>');
    expect(outDocXml).toContain('<wp:wrapTopAndBottom/>');
    expect(outDocXml).toContain('r:embed="rId8"');
    expect(outDocXml).toContain('w:header="708" w:footer="708"');
    expect(outDocXml).toContain('<w:headerReference w:type="default" r:id="rId6"/>');
    expect(outDocXml).toContain('<w:footerReference w:type="default" r:id="rId7"/>');
    expect(outDocXml).toContain('<w:pgBorders w:offsetFrom="page">');
  });

  it('regenerates XML ONLY for edited or newly inserted paragraphs while keeping all untouched blocks and wp:anchor drawings intact', async () => {
    const {buffer: origBuf, parts} = await buildVelavanStylePackage();
    const imported = await importDocxToHtml(origBuf);
    const layout: DocPageLayout = {...DEFAULT_DOC_LAYOUT, ...imported.layoutPatch};

    // Edit 1: change Heading1 text "ANNA UNIVERSITY" -> "ANNA UNIVERSITY CHENNAI"
    // Edit 2: change text in paragraph that shared <w:p> with the anchored logo
    // Edit 3: append a brand-new paragraph with a tab and a brand-new PNG image
    const editedHtml =
      imported.html
        .replace('ANNA UNIVERSITY', 'ANNA UNIVERSITY CHENNAI')
        .replace(
          'ELECTRONICS AND COMMUNICATION ENGINEERING',
          'DEPT OF ELECTRONICS AND COMMUNICATION ENGINEERING',
        ) +
      `<p data-style-id="BodyText" data-tab-stops='[{"posTwips":7200,"posPx":480,"val":"right","leader":"dot"}]'><span data-doc-font="Times New Roman" style="font-size: 13pt; color: #000000">New Line</span><span data-doc-tab="true" data-tab-pos="7200" data-tab-align="right" data-tab-leader="dot">\t</span><span data-doc-font="Times New Roman" style="font-size: 13pt">Page 2</span></p>` +
      `<img src="data:image/png;base64,${BLUE_PNG_B64}" data-width="96" data-height="48" data-align="center" alt="New Stamp" />`;

    const result = await patchOriginalDocxPackage(editedHtml, layout, [], origBuf);
    expect(result).not.toBeNull();
    const {bytes: savedBytes, stats} = result!;

    // 4 untouched body blocks (page break, Heading3, BodyText, EXAMINER tabbed para) stay verbatim!
    expect(stats.preservedBlocks).toBe(4);
    // Generated: edited Heading1, edited ECE text, preserved logo drawing para, new tabbed para, new image para
    expect(stats.generatedBlocks).toBe(5);
    expect(stats.addedMediaCount).toBe(1);
    expect(stats.sectPrPreservedVerbatim).toBe(true);

    const outZip = await JSZip.loadAsync(savedBytes);
    const outDocXml = (await outZip.file('word/document.xml')?.async('string')) ?? '';

    // Auxiliary parts untouched (styles, numbering, theme, fontTable, settings, headers, footers)
    expect(await outZip.file('word/styles.xml')?.async('string')).toBe(parts['word/styles.xml']);
    expect(await outZip.file('word/numbering.xml')?.async('string')).toBe(parts['word/numbering.xml']);
    expect(await outZip.file('word/header1.xml')?.async('string')).toBe(parts['word/header1.xml']);
    expect(await outZip.file('word/footer1.xml')?.async('string')).toBe(parts['word/footer1.xml']);

    // Untouched blocks kept exact original rsidR and XML
    expect(outDocXml).toContain('w:rsidR="00CC5566"');
    expect(outDocXml).toContain('w:rsidR="00DD7788"');
    expect(outDocXml).toContain('w:rsidR="00EE9900"');
    expect(outDocXml).toContain('w:rsidR="00FF1122"');

    // Edited Heading1 has new text AND preserves pStyle Heading1 + font size + black color
    expect(outDocXml).toContain('ANNA UNIVERSITY CHENNAI');
    expect(outDocXml).toContain('<w:pStyle w:val="Heading1"/>');
    expect(outDocXml).toContain('<w:sz w:val="32"/>');
    expect(outDocXml).toContain('<w:color w:val="000000"/>');

    // Anchored logo drawing was STILL preserved with exact wp:anchor and rId8 even though sibling text changed!
    expect(outDocXml).toContain('DEPT OF ELECTRONICS AND COMMUNICATION ENGINEERING');
    expect(outDocXml).toContain('<wp:anchor');
    expect(outDocXml).toContain('<wp:posOffset>2103786</wp:posOffset>');
    expect(outDocXml).toContain('r:embed="rId8"');

    // Newly added tabbed paragraph has real <w:tabs> + <w:tab/> + Times New Roman 13pt (26 half-pts)
    expect(outDocXml).toContain('<w:tab w:val="right" w:leader="dot" w:pos="7200"/>');
    expect(outDocXml).toContain('<w:sz w:val="26"/>');
    expect(outDocXml).toContain('New Line');

    // Newly added image created word/media/esz_image_1.png and appended rId9 to document.xml.rels
    const newImgB64 = await outZip.file('word/media/esz_image_1.png')?.async('base64');
    expect(newImgB64).toBe(BLUE_PNG_B64);
    const outRels = (await outZip.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
    expect(outRels).toContain('Id="rId8"');
    expect(outRels).toContain('Id="rId9"');
    expect(outRels).toContain('Target="media/esz_image_1.png"');
    expect(outDocXml).toContain('r:embed="rId9"');

    // Re-importing the saved .docx must succeed cleanly with all text and images intact
    const reimported = await importDocxToHtml(toArrayBuffer(savedBytes));
    expect(reimported.html).toContain('ANNA UNIVERSITY CHENNAI');
    expect(reimported.html).toContain('BONAFIDE CERTIFICATE');
    expect(reimported.html).toContain('INTERNAL EXAMINER');
    expect(reimported.stats.images).toBe(2);
  });

  it('losslessly round-trips tests/fixtures/calibri-sample.docx via exportHtmlToDocxBytes', async () => {
    const rawBuf = fs.readFileSync(CALIBRI_FIXTURE_PATH);
    const origAb = toArrayBuffer(rawBuf);
    const origZip = await JSZip.loadAsync(origAb);
    const origStyles = await origZip.file('word/styles.xml')?.async('string');
    const origFooter = await origZip.file('word/footer1.xml')?.async('string');

    const imported = await importDocxToHtml(origAb);
    const layout: DocPageLayout = {...DEFAULT_DOC_LAYOUT, ...imported.layoutPatch};

    const savedBytes = await exportHtmlToDocxBytes(imported.html, layout, [], origAb);
    const savedZip = await JSZip.loadAsync(savedBytes);

    expect(await savedZip.file('word/styles.xml')?.async('string')).toBe(origStyles);
    expect(await savedZip.file('word/footer1.xml')?.async('string')).toBe(origFooter);

    const patchRes = await patchOriginalDocxPackage(imported.html, layout, [], origAb);
    expect(patchRes).not.toBeNull();
    expect(patchRes!.stats.generatedBlocks).toBe(0);
    expect(patchRes!.stats.preservedBlocks).toBeGreaterThan(0);
    expect(patchRes!.stats.sectPrPreservedVerbatim).toBe(true);
  });

  describe.skipIf(!hasVelavanFixture)(
    'VELAVAN_FRONT_PAGE_FOR_INTERSHIP.docx Step C round-trip',
    () => {
      it('round-trips VELAVAN original with zero regenerated blocks and identical auxiliary parts', async () => {
        const rawBuf = fs.readFileSync(VELAVAN_FIXTURE_PATH);
        const origAb = toArrayBuffer(rawBuf);
        const origZip = await JSZip.loadAsync(origAb);

        const imported = await importDocxToHtml(origAb);
        const layout: DocPageLayout = {...DEFAULT_DOC_LAYOUT, ...imported.layoutPatch};

        const patchRes = await patchOriginalDocxPackage(imported.html, layout, [], origAb);
        expect(patchRes).not.toBeNull();
        expect(patchRes!.stats.generatedBlocks).toBe(0);
        expect(patchRes!.stats.sectPrPreservedVerbatim).toBe(true);

        const savedZip = await JSZip.loadAsync(patchRes!.bytes);
        for (const p of Object.keys(origZip.files)) {
          if (p === 'word/document.xml' || origZip.files[p].dir) continue;
          const origData = await origZip.file(p)?.async('uint8array');
          const savedData = await savedZip.file(p)?.async('uint8array');
          expect(savedData).toEqual(origData);
        }
      });

      it.skipIf(!hasLibreOfficeTools)(
        'matches original page count, text positions (within 2pt), and images in LibreOffice',
        async () => {
          const rawBuf = fs.readFileSync(VELAVAN_FIXTURE_PATH);
          const origAb = toArrayBuffer(rawBuf);
          const imported = await importDocxToHtml(origAb);
          const layout: DocPageLayout = {...DEFAULT_DOC_LAYOUT, ...imported.layoutPatch};
          const savedBytes = await exportHtmlToDocxBytes(imported.html, layout, [], origAb);

          const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'esz-lo-step-c-'));
          try {
            const copyDocxPath = path.join(tmpDir, 'copy.docx');
            fs.writeFileSync(copyDocxPath, savedBytes);

            execFileSync(
              'soffice',
              ['--headless', '--convert-to', 'pdf', '--outdir', tmpDir, VELAVAN_FIXTURE_PATH],
              {stdio: 'pipe'},
            );
            execFileSync(
              'soffice',
              ['--headless', '--convert-to', 'pdf', '--outdir', tmpDir, copyDocxPath],
              {stdio: 'pipe'},
            );

            const origPdf = path.join(tmpDir, 'VELAVAN_FRONT_PAGE_FOR_INTERSHIP.pdf');
            const copyPdf = path.join(tmpDir, 'copy.pdf');

            const origBbox = execFileSync('pdftotext', ['-bbox', origPdf, '-'], {encoding: 'utf8'});
            const copyBbox = execFileSync('pdftotext', ['-bbox', copyPdf, '-'], {encoding: 'utf8'});

            const parseWords = (xml: string) =>
              Array.from(
                xml.matchAll(/<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]+)<\/word>/g),
              ).map((m) => ({
                xMin: Number(m[1]),
                yMin: Number(m[2]),
                text: m[5],
              }));

            const origPages = (origBbox.match(/<page\b/g) ?? []).length;
            const copyPages = (copyBbox.match(/<page\b/g) ?? []).length;
            expect(copyPages).toBe(origPages);

            const origWords = parseWords(origBbox);
            const copyWords = parseWords(copyBbox);
            expect(copyWords.length).toBe(origWords.length);
            for (let i = 0; i < origWords.length; i++) {
              expect(copyWords[i].text).toBe(origWords[i].text);
              expect(Math.abs(copyWords[i].xMin - origWords[i].xMin)).toBeLessThanOrEqual(2);
              expect(Math.abs(copyWords[i].yMin - origWords[i].yMin)).toBeLessThanOrEqual(2);
            }

            const origImgs = execFileSync('pdfimages', ['-list', origPdf], {encoding: 'utf8'});
            const copyImgs = execFileSync('pdfimages', ['-list', copyPdf], {encoding: 'utf8'});
            expect(copyImgs.trim().split('\n').length).toBe(origImgs.trim().split('\n').length);
          } finally {
            fs.rmSync(tmpDir, {recursive: true, force: true});
          }
        },
      );
    },
  );
});
