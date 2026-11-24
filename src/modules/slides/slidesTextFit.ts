import {
  SLIDE_HEIGHT,
  SLIDE_INCH_H,
  SLIDE_INCH_W,
  SLIDE_WIDTH,
  type SlideAspectRatio,
  type SlideItem,
  type SlideTheme,
  type TextElement,
} from './slidesTypes';

export interface SlideDimensions {
  aspectRatio: SlideAspectRatio;
  width: number;
  height: number;
  inchW: number;
  inchH: number;
  cssAspect: 'aspect-video' | 'aspect-[4/3]';
}

export function getSlideDimensions(aspectRatio: SlideAspectRatio = '16:9'): SlideDimensions {
  if (aspectRatio === '4:3') {
    return {
      aspectRatio: '4:3',
      width: SLIDE_WIDTH,
      height: 720,
      inchW: 10,
      inchH: 7.5,
      cssAspect: 'aspect-[4/3]',
    };
  }
  return {
    aspectRatio: '16:9',
    width: SLIDE_WIDTH,
    height: SLIDE_HEIGHT,
    inchW: SLIDE_INCH_W,
    inchH: SLIDE_INCH_H,
    cssAspect: 'aspect-video',
  };
}

/**
 * Infers canonical slide aspect ratio ('16:9' vs '4:3') from OOXML <p:sldSz> EMUs.
 */
export function inferAspectRatioFromEmu(
  cxEmu?: number | null,
  cyEmu?: number | null,
  typeAttr?: string,
): SlideAspectRatio {
  if (typeAttr) {
    const lower = typeAttr.toLowerCase();
    if (lower.includes('4x3') || lower === 'letter' || lower === 'a4') {
      return '4:3';
    }
    if (lower.includes('16x9') || lower.includes('16x10')) {
      return '16:9';
    }
  }
  if (cxEmu && cyEmu && cxEmu > 0 && cyEmu > 0) {
    const ratio = cxEmu / cyEmu;
    if (ratio < 1.55) {
      return '4:3';
    }
  }
  return '16:9';
}

export type SlideFontCategory = 'sans' | 'serif' | 'mono';

const SERIF_KEYWORDS = [
  'serif',
  'georgia',
  'times',
  'cambria',
  'garamond',
  'palatino',
  'book antiqua',
  'baskerville',
  'didot',
  'bodoni',
  'charter',
  'merriweather',
  'playfair',
  'constantia',
  'century',
  'minion',
  'lora',
  'crimson',
];

const MONO_KEYWORDS = [
  'mono',
  'courier',
  'consolas',
  'menlo',
  'monaco',
  'fira code',
  'jetbrains',
  'source code',
  'lucida console',
  'inconsolata',
  'dejavu sans mono',
];

export function classifyFontCategory(fontName: string | undefined): SlideFontCategory {
  const lower = (fontName || '').toLowerCase();
  if (!lower) return 'sans';
  if (MONO_KEYWORDS.some((k) => lower.includes(k))) return 'mono';
  // Exclude "sans-serif" from matching "serif"
  const withoutSansSerif = lower.replace(/sans-serif/g, '');
  if (SERIF_KEYWORDS.some((k) => withoutSansSerif.includes(k))) return 'serif';
  return 'sans';
}

export function cleanFontName(raw: string | undefined): string {
  if (!raw) return '';
  return raw
    .split(',')[0]
    .replace(/['";{}]/g, '')
    .trim();
}

export interface ImportedThemeFonts {
  major?: string;
  minor?: string;
}

/**
 * Resolves OOXML font names including theme tokens (+mj-lt / +mn-lt / +mj-ea / +mn-ea)
 * into a clean primary font family name with graceful fallback to deck theme fonts.
 */
export function resolveImportedFontName(
  rawTypeface: string | undefined,
  themeFonts: ImportedThemeFonts,
  role: TextElement['role'] = 'body',
  theme?: Pick<SlideTheme, 'headingFont' | 'bodyFont'>,
): string {
  const trimmed = cleanFontName(rawTypeface);
  const fallbackHeading = theme?.headingFont || 'Plus Jakarta Sans';
  const fallbackBody = theme?.bodyFont || 'DM Sans';

  if (!trimmed) {
    if (role === 'title' || role === 'subtitle') {
      return cleanFontName(themeFonts.major) || fallbackHeading;
    }
    return cleanFontName(themeFonts.minor) || fallbackBody;
  }

  const lower = trimmed.toLowerCase();
  if (lower.startsWith('+mj-')) {
    return cleanFontName(themeFonts.major) || fallbackHeading;
  }
  if (lower.startsWith('+mn-')) {
    return cleanFontName(themeFonts.minor) || fallbackBody;
  }

  return trimmed;
}

function quoteCssFont(name: string): string {
  const clean = name.trim();
  if (!clean) return '';
  const generics = new Set([
    'sans-serif',
    'serif',
    'monospace',
    'system-ui',
    '-apple-system',
    'ui-monospace',
    'ui-serif',
    'ui-sans-serif',
  ]);
  if (generics.has(clean.toLowerCase())) return clean;
  if (/[\s]/.test(clean) && !clean.startsWith("'") && !clean.startsWith('"')) {
    return `'${clean}'`;
  }
  return clean;
}

/**
 * Builds a metric-graceful CSS / Canvas font fallback stack so imported Office or
 * proprietary fonts always render cleanly on every platform without missing glyphs
 * or jarring default-serif substitution.
 */
export function getSlideFontCssStack(fontFamily: string | undefined): string {
  const primary = cleanFontName(fontFamily) || 'DM Sans';
  const category = classifyFontCategory(primary);

  const fallbacks =
    category === 'mono'
      ? ['Consolas', 'Courier New', 'ui-monospace', 'Menlo', 'monospace']
      : category === 'serif'
        ? ['Georgia', 'Cambria', 'Times New Roman', 'serif']
        : [
            'DM Sans',
            'Plus Jakarta Sans',
            'Segoe UI',
            'Arial',
            'Helvetica',
            'system-ui',
            '-apple-system',
            'sans-serif',
          ];

  const seen = new Set<string>();
  const stack: string[] = [];
  for (const entry of [primary, ...fallbacks]) {
    const key = entry.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    stack.push(quoteCssFont(entry));
  }
  return stack.join(', ');
}

export function formatSlideTextForFit(
  text: string,
  listStyle?: TextElement['listStyle'],
): string {
  const lines = (text || '').split('\n');
  if (!listStyle || listStyle === 'none') return text || '';
  if (listStyle === 'bullet') {
    return lines
      .map((line) => {
        const trimmed = line.trim();
        if (!trimmed) return '';
        return trimmed.startsWith('•') || trimmed.startsWith('-')
          ? trimmed
          : `• ${trimmed}`;
      })
      .join('\n');
  }
  let counter = 1;
  return lines
    .map((line) => {
      const trimmed = line.trim();
      if (!trimmed) return '';
      const stripped = trimmed.replace(/^\d+[.)]\s*/, '');
      const formatted = `${counter}. ${stripped}`;
      counter += 1;
      return formatted;
    })
    .join('\n');
}

export interface TextFitInput {
  text: string;
  listStyle?: TextElement['listStyle'];
  width: number;
  height: number;
  fontSize: number;
  lineHeight?: number;
  padding?: number;
  bold?: boolean;
  italic?: boolean;
  fontFamily?: string;
  allowAutoShrink?: boolean;
  minFontSize?: number;
  measureText?: (sample: string, fontSpec: string) => number;
}

export interface TextFitResult {
  fittedFontSize: number;
  fittedLineHeight: number;
  wasAutoShrunk: boolean;
  stillOverflows: boolean;
  lineCount: number;
  requiredHeight: number;
  availableHeight: number;
  recommendedHeight: number;
}

let cachedMeasureCtx: CanvasRenderingContext2D | null | undefined;

function getBrowserMeasureContext(): CanvasRenderingContext2D | null {
  if (cachedMeasureCtx !== undefined) return cachedMeasureCtx;
  if (typeof document === 'undefined') {
    cachedMeasureCtx = null;
    return null;
  }
  try {
    const canvas = document.createElement('canvas');
    cachedMeasureCtx = canvas.getContext('2d');
  } catch {
    cachedMeasureCtx = null;
  }
  return cachedMeasureCtx;
}

function estimateStringWidthFallback(
  str: string,
  fontSize: number,
  bold: boolean,
  category: SlideFontCategory,
): number {
  const baseFactor =
    category === 'mono' ? 0.62 : category === 'serif' ? 0.54 : 0.56;
  const weightFactor = bold ? 1.07 : 1.0;
  let units = 0;
  for (let i = 0; i < str.length; i += 1) {
    const ch = str[i];
    if (ch === ' ') {
      units += 0.34;
    } else if ('ilI1|.,!;:'.includes(ch)) {
      units += category === 'mono' ? 1.0 : 0.45;
    } else if ('mwMW@#%&'.includes(ch)) {
      units += category === 'mono' ? 1.0 : 1.38;
    } else if (ch >= 'A' && ch <= 'Z') {
      units += category === 'mono' ? 1.0 : 1.16;
    } else {
      units += 1.0;
    }
  }
  // 1.05 safety margin absorbs cross-platform fallback font width drift
  return units * fontSize * baseFactor * weightFactor * 1.05;
}

function measureLineCount(
  formattedText: string,
  availW: number,
  fontSize: number,
  bold: boolean,
  italic: boolean,
  fontStack: string,
  category: SlideFontCategory,
  customMeasure?: (sample: string, fontSpec: string) => number,
): number {
  const ctx = customMeasure ? null : getBrowserMeasureContext();
  const fontSpec = `${italic ? 'italic ' : ''}${bold ? '700 ' : '400 '}${fontSize}px ${fontStack}`;
  if (ctx) {
    ctx.font = fontSpec;
  }

  const measure = (sample: string): number => {
    if (!sample) return 0;
    if (customMeasure) return customMeasure(sample, fontSpec);
    if (ctx) return ctx.measureText(sample).width * 1.03;
    return estimateStringWidthFallback(sample, fontSize, bold, category);
  };

  const paragraphs = formattedText.split('\n');
  let totalLines = 0;

  for (const para of paragraphs) {
    if (!para.trim()) {
      totalLines += 1;
      continue;
    }
    const words = para.split(/\s+/);
    let currentLine = '';
    for (const word of words) {
      const candidate = currentLine ? `${currentLine} ${word}` : word;
      if (measure(candidate) <= availW) {
        currentLine = candidate;
      } else {
        if (currentLine) {
          totalLines += 1;
          currentLine = '';
        }
        // Single very long unbroken token wraps across multiple lines
        const wordW = measure(word);
        if (wordW > availW) {
          const extraLines = Math.max(1, Math.ceil(wordW / availW));
          totalLines += extraLines - 1;
          currentLine = word;
        } else {
          currentLine = word;
        }
      }
    }
    if (currentLine) {
      totalLines += 1;
    }
  }

  return Math.max(1, totalLines);
}

/**
 * Measures whether a slide text box fits cleanly within its bounding box and,
 * when `allowAutoShrink` is enabled, progressively shrinks fontSize (and gently
 * tightens lineHeight) so imported `.pptx` text never clips or overflows.
 */
export function fitTextToSlideBox(input: TextFitInput): TextFitResult {
  const padding = Math.max(0, input.padding ?? 4);
  const availW = Math.max(20, input.width - padding * 2);
  const availH = Math.max(12, input.height - padding * 2);
  const baseFontSize = Math.max(8, Math.round(input.fontSize || 18));
  const baseLineHeight = Math.max(1.05, input.lineHeight ?? 1.35);
  const allowAutoShrink = input.allowAutoShrink !== false;
  const minFontSize = Math.max(8, Math.min(baseFontSize, input.minFontSize ?? 9));

  const formatted = formatSlideTextForFit(input.text, input.listStyle);
  if (!formatted.trim()) {
    return {
      fittedFontSize: baseFontSize,
      fittedLineHeight: baseLineHeight,
      wasAutoShrunk: false,
      stillOverflows: false,
      lineCount: 1,
      requiredHeight: Math.ceil(baseFontSize * baseLineHeight),
      availableHeight: availH,
      recommendedHeight: input.height,
    };
  }

  const fontStack = getSlideFontCssStack(input.fontFamily);
  const category = classifyFontCategory(input.fontFamily);
  const bold = Boolean(input.bold);
  const italic = Boolean(input.italic);

  const evaluate = (size: number, lh: number) => {
    const lines = measureLineCount(
      formatted,
      availW,
      size,
      bold,
      italic,
      fontStack,
      category,
      input.measureText,
    );
    const reqH = Math.ceil(lines * size * lh);
    return {lines, reqH};
  };

  // First check original font size & lineHeight
  const initial = evaluate(baseFontSize, baseLineHeight);
  if (initial.reqH <= availH + 2) {
    return {
      fittedFontSize: baseFontSize,
      fittedLineHeight: baseLineHeight,
      wasAutoShrunk: false,
      stillOverflows: false,
      lineCount: initial.lines,
      requiredHeight: initial.reqH,
      availableHeight: availH,
      recommendedHeight: Math.max(input.height, initial.reqH + padding * 2 + 4),
    };
  }

  if (!allowAutoShrink) {
    return {
      fittedFontSize: baseFontSize,
      fittedLineHeight: baseLineHeight,
      wasAutoShrunk: false,
      stillOverflows: true,
      lineCount: initial.lines,
      requiredHeight: initial.reqH,
      availableHeight: availH,
      recommendedHeight: Math.ceil(initial.reqH + padding * 2 + 6),
    };
  }

  // First try gently tightening lineHeight at original size if very close
  const tightLh = Math.min(baseLineHeight, 1.18);
  if (tightLh < baseLineHeight) {
    const tightCheck = evaluate(baseFontSize, tightLh);
    if (tightCheck.reqH <= availH + 2) {
      return {
        fittedFontSize: baseFontSize,
        fittedLineHeight: tightLh,
        wasAutoShrunk: true,
        stillOverflows: false,
        lineCount: tightCheck.lines,
        requiredHeight: tightCheck.reqH,
        availableHeight: availH,
        recommendedHeight: Math.max(input.height, tightCheck.reqH + padding * 2 + 4),
      };
    }
  }

  // Iteratively shrink fontSize down to minFontSize
  let bestSize = baseFontSize;
  let bestLh = tightLh;
  let bestEval = evaluate(bestSize, bestLh);

  for (let size = baseFontSize - 1; size >= minFontSize; size -= 1) {
    const candidateLh = size <= 12 ? Math.min(baseLineHeight, 1.14) : tightLh;
    const candidate = evaluate(size, candidateLh);
    bestSize = size;
    bestLh = candidateLh;
    bestEval = candidate;
    if (candidate.reqH <= availH + 2) {
      return {
        fittedFontSize: bestSize,
        fittedLineHeight: bestLh,
        wasAutoShrunk: true,
        stillOverflows: false,
        lineCount: bestEval.lines,
        requiredHeight: bestEval.reqH,
        availableHeight: availH,
        recommendedHeight: Math.max(input.height, bestEval.reqH + padding * 2 + 4),
      };
    }
  }

  // Still overflows even at minFontSize
  return {
    fittedFontSize: bestSize,
    fittedLineHeight: bestLh,
    wasAutoShrunk: bestSize < baseFontSize || bestLh < baseLineHeight,
    stillOverflows: bestEval.reqH > availH + 2,
    lineCount: bestEval.lines,
    requiredHeight: bestEval.reqH,
    availableHeight: availH,
    recommendedHeight: Math.ceil(bestEval.reqH + padding * 2 + 6),
  };
}

export interface SlideTextFitSummary {
  overflowIds: string[];
  autoShrunkIds: string[];
  byElementId: Record<string, TextFitResult>;
}

export function analyzeSlideTextFit(slide: SlideItem): SlideTextFitSummary {
  const overflowIds: string[] = [];
  const autoShrunkIds: string[] = [];
  const byElementId: Record<string, TextFitResult> = {};

  for (const el of slide.elements) {
    if (el.type !== 'text') continue;
    const fit = fitTextToSlideBox({
      text: el.text,
      listStyle: el.listStyle,
      width: el.width,
      height: el.height,
      fontSize: el.fontSize,
      lineHeight: el.lineHeight ?? 1.35,
      padding: el.padding ?? 4,
      bold: el.bold,
      italic: el.italic,
      fontFamily: el.fontFamily,
      allowAutoShrink: el.autoShrink !== false,
    });
    byElementId[el.id] = fit;
    if (fit.stillOverflows || el.hasOverflow) {
      overflowIds.push(el.id);
    }
    if (fit.wasAutoShrunk || (el.originalFontSize && el.originalFontSize > el.fontSize)) {
      autoShrunkIds.push(el.id);
    }
  }

  return {overflowIds, autoShrunkIds, byElementId};
}
