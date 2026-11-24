export const SLIDE_WIDTH = 960;
export const SLIDE_HEIGHT = 540;
export const SLIDE_INCH_W = 10;
export const SLIDE_INCH_H = 5.625;

export type SlideAspectRatio = '16:9' | '4:3';

export type SlideLayoutKind =
  | 'title'
  | 'title-content'
  | 'two-column'
  | 'section'
  | 'comparison'
  | 'image-caption'
  | 'blank';

export type SlideTransitionKind = 'none' | 'fade' | 'slide' | 'zoom';

export type ShapeKind =
  | 'rect'
  | 'rounded-rect'
  | 'ellipse'
  | 'triangle'
  | 'arrow-right'
  | 'diamond'
  | 'line';

export type SlideChartKind = 'bar' | 'line' | 'pie';

export interface SlideElementBase {
  id: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  opacity?: number;
}

export interface TextElement extends SlideElementBase {
  type: 'text';
  text: string;
  fontFamily: string;
  fontSize: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: 'left' | 'center' | 'right';
  color: string;
  fill?: string;
  padding?: number;
  lineHeight?: number;
  listStyle?: 'none' | 'bullet' | 'numbered';
  role?: 'title' | 'subtitle' | 'body' | 'caption';
  autoShrink?: boolean;
  originalFontSize?: number;
  hasOverflow?: boolean;
}

export interface ShapeElement extends SlideElementBase {
  type: 'shape';
  shapeKind: ShapeKind;
  fill: string;
  stroke: string;
  strokeWidth: number;
  cornerRadius?: number;
  label?: string;
  labelColor?: string;
}

export interface ImageElement extends SlideElementBase {
  type: 'image';
  src: string;
  alt: string;
  cornerRadius?: number;
}

export interface TableElement extends SlideElementBase {
  type: 'table';
  rows: string[][];
  headerBg: string;
  headerColor: string;
  cellBg: string;
  cellColor: string;
  borderColor: string;
  fontSize: number;
}

export interface ChartElement extends SlideElementBase {
  type: 'chart';
  chartKind: SlideChartKind;
  title: string;
  labels: string[];
  values: number[];
  colors: string[];
}

export type SlideElement =
  | TextElement
  | ShapeElement
  | ImageElement
  | TableElement
  | ChartElement;

export interface SlideItem {
  id: string;
  layout: SlideLayoutKind;
  background: string;
  transition: SlideTransitionKind;
  notes: string;
  elements: SlideElement[];
}

export interface SlideMasterSettings {
  footerText: string;
  showSlideNumbers: boolean;
  showAccentBar: boolean;
}

export interface PresentationDeck {
  title: string;
  themeId: string;
  aspectRatio: SlideAspectRatio;
  master: SlideMasterSettings;
  slides: SlideItem[];
  activeSlideId: string;
}

export interface SlideTheme {
  id: string;
  name: string;
  description: string;
  background: string;
  surface: string;
  primary: string;
  accent: string;
  headingColor: string;
  bodyColor: string;
  mutedColor: string;
  headingFont: string;
  bodyFont: string;
  chartPalette: string[];
}

export interface DeckTemplate {
  id: string;
  name: string;
  category: string;
  description: string;
  themeId: string;
  buildSlides: (theme: SlideTheme) => SlideItem[];
}

export interface PersistedSlidesMeta {
  version: 1;
  deck: PresentationDeck;
}

export type SlideRibbonTab = 'home' | 'insert' | 'design' | 'transitions' | 'slideshow';

export type SlideDialogKind =
  | null
  | 'templates'
  | 'outline'
  | 'chart'
  | 'table'
  | 'export'
  | 'master';

export const SLIDE_FONT_FAMILIES = [
  'Plus Jakarta Sans',
  'DM Sans',
  'Georgia',
  'Trebuchet MS',
  'Verdana',
  'Courier New',
];

export const SLIDE_FONT_SIZES = [12, 14, 16, 18, 20, 24, 28, 32, 36, 42, 48, 56, 64];
