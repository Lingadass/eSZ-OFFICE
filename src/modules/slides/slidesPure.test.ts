import {test} from 'vitest';
import assert from 'node:assert/strict';
import {
  alignElementOnSlide,
  applyThemeToDeck,
  createDeckFromTemplate,
  createNewElement,
  deckToPdfBytes,
  deckToPptxBytes,
  duplicateSlideItem,
  getSlideTheme,
  moveElementLayer,
  parseOutlineToSlides,
  pptxBytesToDeck,
  reorderSlides,
} from './slidesPure';
import {SLIDE_WIDTH, type TextElement} from './slidesTypes';

test('parseOutlineToSlides converts markdown headings, bullets, notes, and starter syntax into slides', () => {
  const md = [
    '# Launch Readiness',
    '## Q4 Offline Productivity Release',
    '> Welcome attendees and state goals',
    '# Architecture Pillars',
    '- IndexedDB + OPFS local vault',
    '- Zero cloud uploads',
    '- Service Worker offline precache',
    '> Mention zero backend costs',
  ].join('\n');

  const slides = parseOutlineToSlides(md, 'executive-slate');
  assert.equal(slides.length, 2);
  assert.equal(slides[0].layout, 'title');
  assert.equal(slides[0].notes, 'Welcome attendees and state goals');
  assert.equal(slides[1].layout, 'title-content');
  assert.equal(slides[1].notes, 'Mention zero backend costs');

  const starterSeed =
    'Slide 1: eSZ OFFICE Roadmap\nSlide 2: Offline-First Architecture\nSlide 3: PDF, Docs, Sheets & Slides';
  const starterSlides = parseOutlineToSlides(starterSeed);
  assert.equal(starterSlides.length, 3);
  const firstHeading = starterSlides[0].elements.find((e): e is TextElement => e.type === 'text');
  assert.equal(firstHeading?.text, 'eSZ OFFICE Roadmap');
});

test('theme switching, slide reorder, duplication, alignment, and z-order layer helpers work cleanly', () => {
  const deck = createDeckFromTemplate('product-roadmap', 'Q4 Strategy');
  assert.equal(deck.slides.length, 4);

  const themed = applyThemeToDeck(deck, 'midnight-keynote');
  assert.equal(themed.themeId, 'midnight-keynote');
  assert.equal(themed.slides[0].background, '#0F172A');

  const reordered = reorderSlides(themed.slides, 0, 2);
  assert.equal(reordered[2].id, themed.slides[0].id);

  const dup = duplicateSlideItem(themed.slides[1]);
  assert.notEqual(dup.id, themed.slides[1].id);
  assert.equal(dup.elements.length, themed.slides[1].elements.length);
  assert.notEqual(dup.elements[0].id, themed.slides[1].elements[0].id);

  const theme = getSlideTheme('executive-slate');
  const box = createNewElement('shape', theme, {shapeKind: 'rect'});
  const centered = alignElementOnSlide(box, 'center');
  assert.equal(centered.x, Math.round((SLIDE_WIDTH - box.width) / 2));

  const layerEls = [
    createNewElement('shape', theme),
    createNewElement('text', theme),
    createNewElement('chart', theme),
  ];
  const brought = moveElementLayer(layerEls, layerEls[0].id, 'front');
  assert.equal(brought[2].id, layerEls[0].id);
  const sent = moveElementLayer(brought, layerEls[0].id, 'back');
  assert.equal(sent[0].id, layerEls[0].id);
});

test('deckToPptxBytes and pptxBytesToDeck round-trip slides, text, tables, and speaker notes', async () => {
  const deck = createDeckFromTemplate('quarterly-review', 'Q4 Business Review');
  const pptxBytes = await deckToPptxBytes(deck);
  assert.ok(pptxBytes.length > 2000, 'PPTX binary should be non-trivial ZIP archive');
  // Check PK zip magic header
  assert.equal(pptxBytes[0], 0x50);
  assert.equal(pptxBytes[1], 0x4b);

  const imported = await pptxBytesToDeck(pptxBytes, 'Q4 Business Review.pptx');
  assert.equal(imported.slides.length, deck.slides.length);

  const slide1Texts = imported.slides[0].elements
    .filter((e): e is TextElement => e.type === 'text')
    .map((e) => e.text)
    .join(' ');
  assert.ok(slide1Texts.includes('Quarterly Business Review'));
  assert.ok(imported.slides[0].notes.includes('agenda'));

  const slide2Table = imported.slides[1].elements.find((e) => e.type === 'table');
  assert.ok(slide2Table, 'Table should survive PPTX export/import round-trip');
  if (slide2Table && slide2Table.type === 'table') {
    assert.equal(slide2Table.rows[0][0], 'Capability');
  }
});

test('deckToPdfBytes emits a valid multi-page widescreen PDF document', async () => {
  const deck = createDeckFromTemplate('product-roadmap', 'Roadmap PDF');
  const pdfBytes = await deckToPdfBytes(deck);
  assert.ok(pdfBytes.length > 1000);
  const header = new TextDecoder().decode(pdfBytes.slice(0, 5));
  assert.equal(header, '%PDF-');
});

test('pptxBytesToDeck preserves 4:3 aspect ratio, resolves OOXML theme fonts, and auto-shrinks overflowing imported text', async () => {
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();

  // Standard 4:3 presentation.xml (10" x 7.5" = 9144000 x 6858000 EMUs)
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
    </p:presentation>`,
  );

  // Theme with Cambria heading and Calibri body
  zip.file(
    'ppt/theme/theme1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:fontScheme name="Office">
          <a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
          <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
        </a:fontScheme>
      </a:themeElements>
    </a:theme>`,
  );

  // Slide 1 with a title using +mj-lt and a compact text box stuffed with a long 28pt paragraph
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:spPr>
              <a:xfrm><a:off x="457200" y="457200"/><a:ext cx="8229600" cy="914400"/></a:xfrm>
              <a:prstGeom prst="rect"/>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:p><a:r><a:rPr sz="3200" b="1"><a:solidFill><a:srgbClr val="1E3A8A"/></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>Classic 4:3 Deck</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:spPr>
              <a:xfrm><a:off x="457200" y="1828800"/><a:ext cx="3657600" cy="1097280"/></a:xfrm>
              <a:prstGeom prst="rect"/>
            </p:spPr>
            <p:txBody>
              <a:bodyPr lIns="45720" tIns="45720" rIns="45720" bIns="45720"/>
              <a:p><a:r><a:rPr sz="2400"><a:latin typeface="+mn-lt"/></a:rPr><a:t>This imported 4:3 PowerPoint body box intentionally carries a long narrative paragraph so the importer must auto-shrink the font size cleanly rather than letting text overflow or clip outside the box boundaries.</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>`,
  );

  zip.file(
    'ppt/notesSlides/notesSlide1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:sp><p:txBody>
        <a:p><a:r><a:t>Speaker note for 4:3 slide</a:t></a:r></a:p>
      </p:txBody></p:sp></p:spTree></p:cSld>
    </p:notes>`,
  );

  const bytes = await zip.generateAsync({type: 'uint8array'});
  const deck = await pptxBytesToDeck(bytes, 'Legacy43.pptx');

  assert.equal(deck.aspectRatio, '4:3');
  assert.equal(deck.slides.length, 1);
  assert.equal(deck.slides[0].notes, 'Speaker note for 4:3 slide');

  const texts = deck.slides[0].elements.filter((e): e is TextElement => e.type === 'text');
  assert.equal(texts.length, 2);
  assert.equal(texts[0].fontFamily, 'Cambria');
  assert.equal(texts[0].color, '#1E3A8A');
  assert.equal(texts[0].fill, undefined, 'transparent text box must not steal text color as background fill');
  assert.equal(texts[1].fontFamily, 'Calibri');
  assert.ok(texts[1].originalFontSize && texts[1].fontSize < texts[1].originalFontSize, 'overflowing box should auto-shrink');
  assert.ok(!texts[1].hasOverflow, 'auto-shrunk text should fit cleanly with zero remaining overflow');

  // Round-trip 4:3 back through deckToPptxBytes -> pptxBytesToDeck
  const exported43 = await deckToPptxBytes(deck);
  const reimported43 = await pptxBytesToDeck(exported43, 'Legacy43RoundTrip.pptx');
  assert.equal(reimported43.aspectRatio, '4:3');
});

test('pptxBytesToDeck maps modern 13.333in x 7.5in widescreen EMUs without cropping right-edge elements', async () => {
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();

  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:sldSz cx="12192000" cy="6858000"/>
    </p:presentation>`,
  );

  // Element positioned in the right half of a 13.333" slide (x = 9.0", w = 3.5")
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:spPr>
              <a:xfrm><a:off x="8229600" y="685800"/><a:ext cx="3200400" cy="1371600"/></a:xfrm>
              <a:prstGeom prst="rect"/>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Segoe UI"/></a:rPr><a:t>Right-side KPI card</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>`,
  );

  const bytes = await zip.generateAsync({type: 'uint8array'});
  const deck = await pptxBytesToDeck(bytes, 'Modern169.pptx');

  assert.equal(deck.aspectRatio, '16:9');
  const el = deck.slides[0].elements[0];
  assert.ok(el);
  assert.ok(el.x >= 620 && el.x <= 680, `expected x near 648px on 960px slide, got ${el.x}`);
  assert.ok(el.x + el.width <= 960, 'element must stay inside 960px slide width');
});

