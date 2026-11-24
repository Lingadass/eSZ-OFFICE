export type PdfLayoutMode = 'single' | 'continuous' | 'two-page';
export type PdfZoomMode = 'custom' | 'fit-width' | 'fit-page';

export type PdfAnnotationTool =
  | 'select'
  | 'highlight'
  | 'underline'
  | 'strikethrough'
  | 'sticky'
  | 'freehand'
  | 'rect'
  | 'ellipse'
  | 'arrow'
  | 'text'
  | 'stamp'
  | 'image'
  | 'signature'
  | 'redact';

export interface PdfPoint {
  x: number; // Normalized 0..1 relative to page width
  y: number; // Normalized 0..1 relative to page height (top-down)
}

export interface PdfAnnotationItem {
  id: string;
  pageIndex: number; // 0-based
  tool: Exclude<PdfAnnotationTool, 'select'>;
  x: number; // Normalized 0..1
  y: number; // Normalized 0..1
  width: number; // Normalized 0..1
  height: number; // Normalized 0..1
  color: string; // Hex e.g. #FACC15
  opacity: number; // 0..1
  strokeWidth: number; // px at scale 1
  text?: string;
  points?: PdfPoint[]; // For freehand
  imageDataUrl?: string; // For embedded image or signature
}

export type PdfWatermarkPosition =
  | 'center-diagonal'
  | 'center-horizontal'
  | 'top-margin'
  | 'bottom-margin'
  | 'bottom-right';

export type PdfPageNumberPosition =
  | 'top-left'
  | 'top-center'
  | 'top-right'
  | 'bottom-left'
  | 'bottom-center'
  | 'bottom-right';

export type PdfPageNumberFormat = 'number' | 'page-of-total' | 'slash';

export interface PdfWatermarkConfig {
  enabled: boolean;
  text: string;
  opacity: number;
  fontSize: number;
  rotationDeg: number;
  color: string;
  position?: PdfWatermarkPosition;
  skipFirstPage?: boolean;
}

export interface PdfHeaderFooterConfig {
  enabled: boolean;
  headerLeft: string;
  headerCenter: string;
  headerRight: string;
  footerLeft: string;
  footerCenter: string;
  footerRight: string;
  fontSize?: number;
  marginPt?: number;
  skipFirstPage?: boolean;
  ensureSafeMargins?: boolean;
}

export interface PdfPageNumberConfig {
  enabled: boolean;
  position: PdfPageNumberPosition;
  format: PdfPageNumberFormat;
  startNumber: number;
  fontSize?: number;
  marginPt?: number;
  skipFirstPage?: boolean;
  ensureSafeMargins?: boolean;
}

export interface PdfPageTextClearance {
  pageWidth: number;
  pageHeight: number;
  topClearancePt: number;
  bottomClearancePt: number;
  leftClearancePt: number;
  rightClearancePt: number;
  hasText: boolean;
}

export interface PdfMetadataInfo {
  title: string;
  author: string;
  subject: string;
  keywords: string;
  creator: string;
}

export interface PdfFormFieldInfo {
  name: string;
  type: 'text' | 'checkbox' | 'dropdown' | 'radio' | 'unknown';
  value: string | boolean;
  options?: string[];
}
