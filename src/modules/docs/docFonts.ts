/**
 * Metric-compatible offline document typography and font substitution for eSZ OFFICE.
 *
 * All bundled fonts are self-hosted SIL Open Font License 1.1 (.woff2) in /public/fonts
 * and precached by the service worker for 100% offline rendering fidelity:
 * - Times New Roman -> Liberation Serif
 * - Arial           -> Liberation Sans
 * - Calibri         -> Carlito
 * - Cambria         -> Caladea
 * - Courier New     -> Liberation Mono
 *
 * Substitution is strictly render-only: canonical Word font names are preserved in
 * the document model and in exported .docx files, and app shell fonts
 * (Plus Jakarta Sans, DM Sans) are never applied inside documents.
 */

export interface BundledFontFaceSpec {
  family: string;
  weight: 400 | 700;
  style: 'normal' | 'italic';
  url: string;
}

export const FONT_SUBSTITUTION_MAP: Readonly<Record<string, string>> = {
  'Times New Roman': 'Liberation Serif',
  Arial: 'Liberation Sans',
  Calibri: 'Carlito',
  Cambria: 'Caladea',
  'Courier New': 'Liberation Mono',
};

const LOWER_TO_CANONICAL_AND_SUB: Readonly<
  Record<string, {canonical: string; substitute: string; generic: 'serif' | 'sans-serif' | 'monospace'}>
> = {
  'times new roman': {
    canonical: 'Times New Roman',
    substitute: 'Liberation Serif',
    generic: 'serif',
  },
  times: {
    canonical: 'Times New Roman',
    substitute: 'Liberation Serif',
    generic: 'serif',
  },
  arial: {
    canonical: 'Arial',
    substitute: 'Liberation Sans',
    generic: 'sans-serif',
  },
  helvetica: {
    canonical: 'Arial',
    substitute: 'Liberation Sans',
    generic: 'sans-serif',
  },
  calibri: {
    canonical: 'Calibri',
    substitute: 'Carlito',
    generic: 'sans-serif',
  },
  'calibri light': {
    canonical: 'Calibri Light',
    substitute: 'Carlito',
    generic: 'sans-serif',
  },
  cambria: {
    canonical: 'Cambria',
    substitute: 'Caladea',
    generic: 'serif',
  },
  'cambria math': {
    canonical: 'Cambria Math',
    substitute: 'Caladea',
    generic: 'serif',
  },
  'courier new': {
    canonical: 'Courier New',
    substitute: 'Liberation Mono',
    generic: 'monospace',
  },
  courier: {
    canonical: 'Courier New',
    substitute: 'Liberation Mono',
    generic: 'monospace',
  },
};

const SUBSTITUTE_TO_CANONICAL_LOWER: Readonly<Record<string, string>> = {
  'liberation serif': 'times new roman',
  'liberation sans': 'arial',
  carlito: 'calibri',
  caladea: 'cambria',
  'liberation mono': 'courier new',
};

const GENERIC_FAMILIES = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
  'ui-sans-serif',
  'ui-serif',
  'ui-monospace',
  'inherit',
  'initial',
  'unset',
]);

/** App shell UI fonts that must NEVER be applied inside document bodies or exports */
const FORBIDDEN_DOC_FONTS = new Set([
  'plus jakarta sans',
  'dm sans',
  'jetbrains mono',
]);

export const DOCUMENT_FONT_PRESETS: readonly string[] = [
  'Calibri',
  'Times New Roman',
  'Arial',
  'Cambria',
  'Courier New',
  'Georgia',
  'Verdana',
];

export const BUNDLED_FONT_FACES: readonly BundledFontFaceSpec[] = [
  // Liberation Serif (Times New Roman metric equivalent — SIL OFL 1.1)
  {family: 'Liberation Serif', weight: 400, style: 'normal', url: '/fonts/liberation-serif-400-normal.woff2'},
  {family: 'Liberation Serif', weight: 700, style: 'normal', url: '/fonts/liberation-serif-700-normal.woff2'},
  {family: 'Liberation Serif', weight: 400, style: 'italic', url: '/fonts/liberation-serif-400-italic.woff2'},
  {family: 'Liberation Serif', weight: 700, style: 'italic', url: '/fonts/liberation-serif-700-italic.woff2'},
  // Liberation Sans (Arial metric equivalent — SIL OFL 1.1)
  {family: 'Liberation Sans', weight: 400, style: 'normal', url: '/fonts/liberation-sans-400-normal.woff2'},
  {family: 'Liberation Sans', weight: 700, style: 'normal', url: '/fonts/liberation-sans-700-normal.woff2'},
  {family: 'Liberation Sans', weight: 400, style: 'italic', url: '/fonts/liberation-sans-400-italic.woff2'},
  {family: 'Liberation Sans', weight: 700, style: 'italic', url: '/fonts/liberation-sans-700-italic.woff2'},
  // Liberation Mono (Courier New metric equivalent — SIL OFL 1.1)
  {family: 'Liberation Mono', weight: 400, style: 'normal', url: '/fonts/liberation-mono-400-normal.woff2'},
  {family: 'Liberation Mono', weight: 700, style: 'normal', url: '/fonts/liberation-mono-700-normal.woff2'},
  {family: 'Liberation Mono', weight: 400, style: 'italic', url: '/fonts/liberation-mono-400-italic.woff2'},
  {family: 'Liberation Mono', weight: 700, style: 'italic', url: '/fonts/liberation-mono-700-italic.woff2'},
  // Carlito (Calibri metric equivalent — SIL OFL 1.1)
  {family: 'Carlito', weight: 400, style: 'normal', url: '/fonts/carlito-400-normal.woff2'},
  {family: 'Carlito', weight: 700, style: 'normal', url: '/fonts/carlito-700-normal.woff2'},
  {family: 'Carlito', weight: 400, style: 'italic', url: '/fonts/carlito-400-italic.woff2'},
  {family: 'Carlito', weight: 700, style: 'italic', url: '/fonts/carlito-700-italic.woff2'},
  // Caladea (Cambria metric equivalent — SIL OFL 1.1)
  {family: 'Caladea', weight: 400, style: 'normal', url: '/fonts/caladea-400-normal.woff2'},
  {family: 'Caladea', weight: 700, style: 'normal', url: '/fonts/caladea-700-normal.woff2'},
  {family: 'Caladea', weight: 400, style: 'italic', url: '/fonts/caladea-400-italic.woff2'},
  {family: 'Caladea', weight: 700, style: 'italic', url: '/fonts/caladea-700-italic.woff2'},
];

function stripQuotes(name: string): string {
  return name.trim().replace(/^["']+|["']+$/g, '').trim();
}

function quoteIfNeeded(name: string): string {
  if (GENERIC_FAMILIES.has(name.toLowerCase())) return name.toLowerCase();
  return `"${name.replace(/"/g, '')}"`;
}

function splitFontList(raw: string): string[] {
  const parts: string[] = [];
  let current = '';
  let quoteChar: '"' | "'" | null = null;
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (quoteChar) {
      if (ch === quoteChar) quoteChar = null;
      current += ch;
    } else if (ch === '"' || ch === "'") {
      quoteChar = ch;
      current += ch;
    } else if (ch === ',') {
      if (current.trim()) parts.push(stripQuotes(current));
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.trim()) parts.push(stripQuotes(current));
  return parts.filter(Boolean);
}

/**
 * Returns the metric-compatible SIL OFL substitute family for a Windows/Office font name,
 * or null if no substitution is defined.
 */
export function getSubstituteFontFamily(fontName: string | null | undefined): string | null {
  if (!fontName) return null;
  const clean = stripQuotes(fontName).toLowerCase();
  return LOWER_TO_CANONICAL_AND_SUB[clean]?.substitute ?? null;
}

/**
 * Returns true if the font name is an app shell UI font (Plus Jakarta Sans, DM Sans, JetBrains Mono)
 * that must never leak into document styling or exported .docx files.
 */
export function isAppShellUiFont(fontName: string | null | undefined): boolean {
  if (!fontName) return false;
  return FORBIDDEN_DOC_FONTS.has(stripQuotes(fontName).toLowerCase());
}

/**
 * Extracts the original canonical Word font name from a single name or rendered CSS stack,
 * preserving the original document font in the model and exported .docx.
 *
 * Examples:
 * - '"Carlito", "Calibri", sans-serif' -> 'Calibri'
 * - '"Liberation Serif", "Times New Roman", serif' -> 'Times New Roman'
 * - 'Times New Roman' -> 'Times New Roman'
 * - '"DM Sans", sans-serif' -> null
 */
export function canonicalizeWordFontFamily(fontFamily: string | null | undefined): string | null {
  if (!fontFamily) return null;
  const tokens = splitFontList(fontFamily).filter(
    (t) => !GENERIC_FAMILIES.has(t.toLowerCase()) && !FORBIDDEN_DOC_FONTS.has(t.toLowerCase()),
  );
  if (tokens.length === 0) return null;

  // If the first token is one of our render-only substitutes AND the paired canonical
  // Word font appears right after it in the rendered stack, recover the original Word font.
  const firstLower = tokens[0].toLowerCase();
  const pairedCanonicalLower = SUBSTITUTE_TO_CANONICAL_LOWER[firstLower];
  if (pairedCanonicalLower) {
    const pairedMatch = tokens.find(
      (t) =>
        t.toLowerCase() === pairedCanonicalLower ||
        LOWER_TO_CANONICAL_AND_SUB[t.toLowerCase()]?.substitute.toLowerCase() === firstLower,
    );
    if (pairedMatch) {
      return LOWER_TO_CANONICAL_AND_SUB[pairedMatch.toLowerCase()]?.canonical ?? pairedMatch;
    }
  }

  const mapped = LOWER_TO_CANONICAL_AND_SUB[firstLower];
  return mapped ? mapped.canonical : tokens[0];
}

function guessGenericFallback(fontName: string): 'serif' | 'sans-serif' | 'monospace' {
  const lower = fontName.toLowerCase();
  if (LOWER_TO_CANONICAL_AND_SUB[lower]) {
    return LOWER_TO_CANONICAL_AND_SUB[lower].generic;
  }
  if (/mono|courier|consolas|menlo|fixed/i.test(lower)) return 'monospace';
  if (/serif|times|georgia|garamond|palatino|cambria|caladea|book/i.test(lower)) return 'serif';
  return 'sans-serif';
}

/**
 * Resolves a canonical Word font name or CSS font-family stack into a render-only CSS font stack
 * backed by our self-hosted SIL OFL metric-compatible fonts.
 *
 * - Never includes Plus Jakarta Sans or DM Sans.
 * - Keeps the original Word font name alongside its metric-compatible substitute for render fidelity.
 */
export function resolveRenderFontStack(
  fontFamily: string | null | undefined,
  fallbackCanonical = 'Times New Roman',
): string {
  const rawList = fontFamily ? splitFontList(fontFamily) : [];
  const filtered = rawList.filter((t) => !FORBIDDEN_DOC_FONTS.has(t.toLowerCase()));

  const inputTokens = filtered.length > 0 ? filtered : [fallbackCanonical];
  const out: string[] = [];
  const seen = new Set<string>();
  let detectedGeneric: 'serif' | 'sans-serif' | 'monospace' | null = null;

  const pushUnique = (name: string) => {
    const key = name.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    out.push(quoteIfNeeded(name));
  };

  for (const token of inputTokens) {
    const lower = token.toLowerCase();
    if (GENERIC_FAMILIES.has(lower)) {
      if (lower === 'serif' || lower === 'sans-serif' || lower === 'monospace') {
        detectedGeneric = detectedGeneric ?? lower;
      }
      continue;
    }

    const entry = LOWER_TO_CANONICAL_AND_SUB[lower];
    if (entry) {
      // Place metric-compatible self-hosted SIL OFL font first for deterministic cross-platform offline rendering
      pushUnique(entry.substitute);
      pushUnique(entry.canonical);
      detectedGeneric = detectedGeneric ?? entry.generic;
    } else {
      pushUnique(token);
      if (!detectedGeneric) {
        detectedGeneric = guessGenericFallback(token);
      }
    }
  }

  // Ensure a metric-compatible bundled fallback + generic family are always present
  if (!detectedGeneric) {
    detectedGeneric = guessGenericFallback(fallbackCanonical);
  }
  if (out.length === 0) {
    const fb = LOWER_TO_CANONICAL_AND_SUB[fallbackCanonical.toLowerCase()];
    if (fb) {
      pushUnique(fb.substitute);
      pushUnique(fb.canonical);
    } else {
      pushUnique('Liberation Serif');
    }
  }
  pushUnique(detectedGeneric);

  return out.join(', ');
}

/**
 * Rewrites CSS `font-family:` declarations inside generated preview stylesheets
 * to use our render-only metric-compatible font substitution map.
 */
export function rewriteCssFontFamiliesForRender(cssText: string): string {
  return cssText.replace(
    /(font-family\s*:\s*)([^;}{!]+?)(\s*!important)?(?=\s*[;}\r\n]|$)/gi,
    (_full, prefix: string, value: string, important: string | undefined) => {
      const trimmed = value.trim();
      if (!trimmed || /^(inherit|initial|unset|var\()/i.test(trimmed)) {
        return `${prefix}${trimmed}${important ? ' !important' : ''}`;
      }
      const substituted = resolveRenderFontStack(trimmed);
      return `${prefix}${substituted}${important ? ' !important' : ''}`;
    },
  );
}

/**
 * Applies render-only font substitution to docx-preview's generated <style> container
 * and any inline `font-family` styles inside the rendered preview DOM.
 * Never mutates the original stored .docx bytes.
 */
export function applyRenderFontSubstitutionToPreviewDom(
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement | null,
): void {
  const rewriteStylesIn = (root: HTMLElement | null | undefined) => {
    if (!root) return;
    const styleEls = root.querySelectorAll('style');
    styleEls.forEach((styleEl) => {
      if (styleEl.textContent) {
        styleEl.textContent = rewriteCssFontFamiliesForRender(styleEl.textContent);
      }
    });
  };

  rewriteStylesIn(styleContainer);
  rewriteStylesIn(bodyContainer);

  const styledEls = bodyContainer.querySelectorAll<HTMLElement>('[style*="font-family"]');
  styledEls.forEach((el) => {
    const raw = el.style.fontFamily;
    if (!raw || /^(inherit|initial|unset|var\()/i.test(raw.trim())) return;
    const substituted = resolveRenderFontStack(raw);
    if (substituted !== raw) {
      el.style.fontFamily = substituted;
    }
  });
}
