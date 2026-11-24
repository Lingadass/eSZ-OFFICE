import {resolveRenderFontStack} from './docFonts';

export type DocsRibbonTab =
  | 'home'
  | 'insert'
  | 'page-layout'
  | 'references'
  | 'review'
  | 'view';

export type DocMarginPreset = 'normal' | 'narrow' | 'moderate' | 'wide';
export type DocOrientation = 'portrait' | 'landscape';
export type DocPageSize = 'a4' | 'letter' | 'legal';
export type DocColumns = 1 | 2 | 3;
export type DocThemeId = 'office-classic' | 'slate-editorial' | 'emerald-modern' | 'warm-serif';
export type DocBorderStyle = 'none' | 'box' | 'double' | 'shadow';
export type DocViewMode = 'print' | 'web';

export interface DocPageBorderSide {
  val: string;
  szPt: number;
  color: string;
  spacePt: number;
}

export interface DocPageBorderConfig {
  offsetFrom: 'page' | 'text';
  display: 'allPages' | 'firstPage' | 'notFirstPage';
  top?: DocPageBorderSide;
  right?: DocPageBorderSide;
  bottom?: DocPageBorderSide;
  left?: DocPageBorderSide;
}

export interface DocTabStop {
  posTwips: number;
  posPx: number;
  val: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear';
  leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface DocPageLayout {
  margins: DocMarginPreset;
  orientation: DocOrientation;
  size: DocPageSize;
  columns: DocColumns;
  theme: DocThemeId;
  watermarkText: string;
  watermarkEnabled: boolean;
  pageBorder: DocBorderStyle;
  headerText: string;
  footerText: string;
  showPageNumbers: boolean;
  pageNumberPosition: 'bottom-center' | 'bottom-right' | 'top-right';
  differentFirstPage?: boolean;
  firstPageHeaderText?: string;
  firstPageFooterText?: string;
  /** Canonical Word default body font from word/styles.xml (docDefaults / Normal), never an app UI font */
  docDefaultFont?: string;
  /** Default body font size in pt from word/styles.xml (docDefaults / Normal) */
  docDefaultFontSizePt?: number;
  /** Default body text color from word/styles.xml (docDefaults / Normal) */
  docDefaultColor?: string;
  /** True when document typography/colors come from imported OOXML styles.xml (disables app theme heading colors) */
  hasImportedStyles?: boolean;
  /** Default tab stop interval in CSS px (from word/settings.xml w:defaultTabStop, default 48px = 0.5in) */
  defaultTabStopPx?: number;
  /** Exact imported OOXML page width in CSS px (96dpi), overrides size preset when set */
  customPageWidthPx?: number;
  /** Exact imported OOXML page height in CSS px (96dpi), overrides size preset when set */
  customPageHeightPx?: number;
  /** Exact imported OOXML margins in CSS px (96dpi), overrides margin preset when set */
  customMarginsPx?: {top: number; right: number; bottom: number; left: number};
  /** Exact imported OOXML w:pgBorders configuration */
  pageBorderConfig?: DocPageBorderConfig | null;
}

export interface DocComment {
  id: string;
  author: string;
  quote: string;
  text: string;
  createdAt: number;
  resolved: boolean;
}

export interface DocNoteItem {
  id: string;
  kind: 'footnote' | 'endnote';
  number: number;
  text: string;
}

export interface DocPersistedMeta {
  layout: DocPageLayout;
  comments: DocComment[];
  notes: DocNoteItem[];
}

export interface DocStats {
  words: number;
  characters: number;
  charactersNoSpaces: number;
  paragraphs: number;
  estimatedPages: number;
  readingMinutes: number;
}

export interface DocHeadingItem {
  level: 1 | 2 | 3;
  text: string;
  slug: string;
}

export interface FormatPainterMarks {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color?: string;
  highlight?: string;
  fontFamily?: string;
  fontSize?: string;
}

export const DEFAULT_DOC_LAYOUT: DocPageLayout = {
  margins: 'normal',
  orientation: 'portrait',
  size: 'a4',
  columns: 1,
  theme: 'office-classic',
  watermarkText: 'CONFIDENTIAL',
  watermarkEnabled: false,
  pageBorder: 'none',
  headerText: '',
  footerText: '',
  showPageNumbers: true,
  pageNumberPosition: 'bottom-center',
  differentFirstPage: false,
  firstPageHeaderText: '',
  firstPageFooterText: '',
};

export const DOC_PAGE_DIMENSIONS_PX: Record<DocPageSize, {width: number; height: number; label: string}> = {
  a4: {width: 794, height: 1123, label: 'A4 (210 × 297 mm)'},
  letter: {width: 816, height: 1056, label: 'Letter (8.5 × 11 in)'},
  legal: {width: 816, height: 1344, label: 'Legal (8.5 × 14 in)'},
};

export const DOC_MARGIN_PX: Record<DocMarginPreset, {top: number; right: number; bottom: number; left: number; inches: string}> = {
  normal: {top: 72, right: 72, bottom: 72, left: 72, inches: '1"'},
  narrow: {top: 38, right: 38, bottom: 38, left: 38, inches: '0.5"'},
  moderate: {top: 72, right: 54, bottom: 72, left: 54, inches: '0.75"'},
  wide: {top: 72, right: 112, bottom: 72, left: 112, inches: '1.5"'},
};

export const DOC_THEMES: Record<
  DocThemeId,
  {
    name: string;
    canonicalFontFamily: string;
    fontFamily: string;
    headingColor: string;
    accentColor: string;
    textColor: string;
  }
> = {
  'office-classic': {
    name: 'Office Classic',
    canonicalFontFamily: 'Calibri',
    fontFamily: resolveRenderFontStack('Calibri'),
    headingColor: '#1E3A8A',
    accentColor: '#1D4ED8',
    textColor: '#0F172A',
  },
  'slate-editorial': {
    name: 'Slate Editorial',
    canonicalFontFamily: 'Cambria',
    fontFamily: resolveRenderFontStack('Cambria'),
    headingColor: '#0F172A',
    accentColor: '#334155',
    textColor: '#1E293B',
  },
  'emerald-modern': {
    name: 'Emerald Modern',
    canonicalFontFamily: 'Arial',
    fontFamily: resolveRenderFontStack('Arial'),
    headingColor: '#065F46',
    accentColor: '#059669',
    textColor: '#0F172A',
  },
  'warm-serif': {
    name: 'Warm Academic',
    canonicalFontFamily: 'Times New Roman',
    fontFamily: resolveRenderFontStack('Times New Roman'),
    headingColor: '#7C2D12',
    accentColor: '#B45309',
    textColor: '#1C1917',
  },
};
