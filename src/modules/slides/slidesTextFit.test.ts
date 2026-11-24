import {test} from 'vitest';
import assert from 'node:assert/strict';
import {
  analyzeSlideTextFit,
  classifyFontCategory,
  fitTextToSlideBox,
  getSlideDimensions,
  getSlideFontCssStack,
  inferAspectRatioFromEmu,
  resolveImportedFontName,
} from './slidesTextFit';
import type {SlideItem} from './slidesTypes';

test('getSlideDimensions and inferAspectRatioFromEmu accurately distinguish 16:9 and 4:3 decks', () => {
  const widescreen = getSlideDimensions('16:9');
  assert.equal(widescreen.width, 960);
  assert.equal(widescreen.height, 540);
  assert.equal(widescreen.cssAspect, 'aspect-video');

  const standard = getSlideDimensions('4:3');
  assert.equal(standard.width, 960);
  assert.equal(standard.height, 720);
  assert.equal(standard.cssAspect, 'aspect-[4/3]');

  // Standard 4:3 (10" x 7.5")
  assert.equal(inferAspectRatioFromEmu(9144000, 6858000), '4:3');
  assert.equal(inferAspectRatioFromEmu(9144000, 6858000, 'screen4x3'), '4:3');
  // Classic 16:9 (10" x 5.625")
  assert.equal(inferAspectRatioFromEmu(9144000, 5143500), '16:9');
  // Modern PowerPoint Widescreen 16:9 (13.333" x 7.5")
  assert.equal(inferAspectRatioFromEmu(12192000, 6858000), '16:9');
});

test('resolveImportedFontName and getSlideFontCssStack provide graceful fallbacks for Office and custom fonts', () => {
  const themeFonts = {major: 'Calibri Light', minor: 'Calibri'};
  const theme = {headingFont: 'Plus Jakarta Sans', bodyFont: 'DM Sans'};

  assert.equal(resolveImportedFontName('+mj-lt', themeFonts, 'title', theme), 'Calibri Light');
  assert.equal(resolveImportedFontName('+mn-lt', themeFonts, 'body', theme), 'Calibri');
  assert.equal(resolveImportedFontName(undefined, {}, 'title', theme), 'Plus Jakarta Sans');
  assert.equal(resolveImportedFontName('AcmeCorpDisplay', themeFonts, 'body', theme), 'AcmeCorpDisplay');

  assert.equal(classifyFontCategory('Calibri'), 'sans');
  assert.equal(classifyFontCategory('Cambria'), 'serif');
  assert.equal(classifyFontCategory('Consolas'), 'mono');
  assert.equal(classifyFontCategory('Unknown Sans-Serif'), 'sans');

  const calibriStack = getSlideFontCssStack('Calibri');
  assert.ok(calibriStack.includes('Calibri'));
  assert.ok(calibriStack.includes('sans-serif'));

  const cambriaStack = getSlideFontCssStack('Cambria');
  assert.ok(cambriaStack.includes('Georgia'));
  assert.ok(cambriaStack.endsWith('serif'));

  const consolasStack = getSlideFontCssStack('Consolas');
  assert.ok(consolasStack.includes('Courier New'));
  assert.ok(consolasStack.endsWith('monospace'));
});

test('fitTextToSlideBox preserves comfortable text, auto-shrinks overflowing text, and flags extreme overflow', () => {
  // Comfortable fit: should not shrink
  const shortFit = fitTextToSlideBox({
    text: 'Quarterly Roadmap',
    width: 600,
    height: 80,
    fontSize: 32,
    fontFamily: 'Calibri',
    allowAutoShrink: true,
  });
  assert.equal(shortFit.wasAutoShrunk, false);
  assert.equal(shortFit.stillOverflows, false);
  assert.equal(shortFit.fittedFontSize, 32);

  // Overflowing text that fits once auto-shrunk
  const mediumOverflow = fitTextToSlideBox({
    text: 'This imported PowerPoint callout contains a detailed paragraph explaining regional offline storage guarantees, local encryption workflows, and zero-cloud latency benchmarks across desktop and mobile devices.',
    width: 340,
    height: 110,
    fontSize: 24,
    fontFamily: 'Calibri',
    allowAutoShrink: true,
  });
  assert.equal(mediumOverflow.wasAutoShrunk, true);
  assert.equal(mediumOverflow.stillOverflows, false);
  assert.ok(mediumOverflow.fittedFontSize < 24, 'fontSize should be auto-shrunk');
  assert.ok(
    mediumOverflow.requiredHeight <= mediumOverflow.availableHeight + 2,
    'shrunk text must fit within box height',
  );

  // Huge text inside a tiny box: shrinks to minFontSize and reports stillOverflows warning
  const extremeOverflow = fitTextToSlideBox({
    text: Array.from(
      {length: 14},
      (_, i) => `Line ${i + 1}: Detailed operational compliance audit item requiring review.`,
    ).join('\n'),
    width: 220,
    height: 44,
    fontSize: 20,
    fontFamily: 'Arial',
    allowAutoShrink: true,
  });
  assert.equal(extremeOverflow.wasAutoShrunk, true);
  assert.equal(extremeOverflow.stillOverflows, true);
  assert.equal(extremeOverflow.fittedFontSize, 9);
  assert.ok(extremeOverflow.recommendedHeight > 44);
});

test('analyzeSlideTextFit reports overflowing and auto-shrunk element IDs on a slide', () => {
  const slide: SlideItem = {
    id: 'slide_1',
    layout: 'blank',
    background: '#FFFFFF',
    transition: 'none',
    notes: '',
    elements: [
      {
        id: 'txt_ok',
        type: 'text',
        x: 40,
        y: 30,
        width: 500,
        height: 70,
        rotation: 0,
        text: 'Clean Title',
        fontFamily: 'DM Sans',
        fontSize: 28,
        color: '#0F172A',
      },
      {
        id: 'txt_tiny_overflow',
        type: 'text',
        x: 40,
        y: 120,
        width: 160,
        height: 30,
        rotation: 0,
        text: 'One\nTwo\nThree\nFour\nFive\nSix\nSeven\nEight\nNine\nTen',
        fontFamily: 'Calibri',
        fontSize: 18,
        color: '#0F172A',
        autoShrink: true,
      },
    ],
  };

  const summary = analyzeSlideTextFit(slide);
  assert.deepEqual(summary.overflowIds, ['txt_tiny_overflow']);
  assert.ok(summary.autoShrunkIds.includes('txt_tiny_overflow'));
  assert.ok(!summary.overflowIds.includes('txt_ok'));
});
