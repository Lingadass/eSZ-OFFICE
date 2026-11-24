import {canonicalizeWordFontFamily} from './docFonts';
import {exportHtmlToPdfHeadless} from './domPdfExport';
import {exportHtmlToDocxBytes} from './docxExport';
import {
  DEFAULT_DOC_LAYOUT,
  type DocHeadingItem,
  type DocNoteItem,
  type DocPageLayout,
  type DocStats,
} from './docsTypes';

export function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function decodeHtmlEntities(str: string): string {
  return str
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

export function htmlToPlainText(html: string): string {
  const withNewlines = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|h[1-6]|li|tr|blockquote)>/gi, '\n')
    .replace(/<td[^>]*>/gi, '\t')
    .replace(/<th[^>]*>/gi, '\t')
    .replace(/<[^>]+>/g, '');
  return decodeHtmlEntities(withNewlines)
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function plainTextToHtml(text: string): string {
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  if (lines.length === 0 || (lines.length === 1 && !lines[0].trim())) {
    return '<p></p>';
  }
  return lines
    .map((line) => (line.trim() ? `<p>${escapeHtml(line)}</p>` : '<p></p>'))
    .join('');
}

export function computeDocumentStats(html: string): DocStats {
  const plain = htmlToPlainText(html);
  const words = plain ? plain.split(/\s+/).filter(Boolean).length : 0;
  const characters = plain.length;
  const charactersNoSpaces = plain.replace(/\s+/g, '').length;
  const paragraphs = plain
    ? plain.split(/\n+/).filter((p) => p.trim().length > 0).length
    : 0;
  const estimatedPages = Math.max(1, Math.ceil(words / 420));
  const readingMinutes = Math.max(1, Math.ceil(words / 220));
  return {
    words,
    characters,
    charactersNoSpaces,
    paragraphs,
    estimatedPages,
    readingMinutes,
  };
}

export function extractHeadingsFromHtml(html: string): DocHeadingItem[] {
  const headings: DocHeadingItem[] = [];
  const regex = /<h([1-3])[^>]*>([\s\S]*?)<\/h\1>/gi;
  let match: RegExpExecArray | null;
  let idx = 0;
  while ((match = regex.exec(html)) !== null) {
    const level = Number(match[1]) as 1 | 2 | 3;
    const text = decodeHtmlEntities(match[2].replace(/<[^>]+>/g, '')).trim();
    if (!text) continue;
    idx += 1;
    const slug = `heading-${idx}-${text
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-|-$/g, '')}`;
    headings.push({level, text, slug});
  }
  return headings;
}

export function buildTocHtml(headings: DocHeadingItem[]): string {
  if (headings.length === 0) {
    return `<div data-doc-toc="true"><p><strong>Table of Contents</strong></p><p>Add Heading 1, Heading 2, or Heading 3 styles to populate entries automatically.</p></div>`;
  }
  const items = headings
    .map((h) => {
      const indent = (h.level - 1) * 18;
      return `<p style="margin-left:${indent}px">${escapeHtml(h.text)}</p>`;
    })
    .join('');
  return `<div data-doc-toc="true"><p><strong>Table of Contents</strong></p>${items}</div>`;
}

export function markdownToHtml(md: string): string {
  const lines = md.replace(/\r\n/g, '\n').split('\n');
  const out: string[] = [];
  let inUl = false;
  let inOl = false;

  const closeLists = () => {
    if (inUl) {
      out.push('</ul>');
      inUl = false;
    }
    if (inOl) {
      out.push('</ol>');
      inOl = false;
    }
  };

  const inlineFormat = (raw: string): string => {
    let s = escapeHtml(raw);
    s = s.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    s = s.replace(/__(.+?)__/g, '<u>$1</u>');
    s = s.replace(/\*(.+?)\*/g, '<em>$1</em>');
    s = s.replace(/~~(.+?)~~/g, '<s>$1</s>');
    s = s.replace(/`([^`]+)`/g, '<code>$1</code>');
    s = s.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2">$1</a>');
    return s;
  };

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      closeLists();
      continue;
    }
    if (/^---+$/.test(trimmed)) {
      closeLists();
      out.push('<hr />');
      continue;
    }
    const hMatch = trimmed.match(/^(#{1,4})\s+(.+)$/);
    if (hMatch) {
      closeLists();
      const lvl = hMatch[1].length;
      out.push(`<h${lvl}>${inlineFormat(hMatch[2])}</h${lvl}>`);
      continue;
    }
    const ulMatch = trimmed.match(/^[-*+]\s+(.+)$/);
    if (ulMatch) {
      if (inOl) {
        out.push('</ol>');
        inOl = false;
      }
      if (!inUl) {
        out.push('<ul>');
        inUl = true;
      }
      out.push(`<li><p>${inlineFormat(ulMatch[1])}</p></li>`);
      continue;
    }
    const olMatch = trimmed.match(/^\d+\.\s+(.+)$/);
    if (olMatch) {
      if (inUl) {
        out.push('</ul>');
        inUl = false;
      }
      if (!inOl) {
        out.push('<ol>');
        inOl = true;
      }
      out.push(`<li><p>${inlineFormat(olMatch[1])}</p></li>`);
      continue;
    }
    const quoteMatch = trimmed.match(/^>\s+(.+)$/);
    if (quoteMatch) {
      closeLists();
      out.push(`<blockquote><p>${inlineFormat(quoteMatch[1])}</p></blockquote>`);
      continue;
    }
    closeLists();
    out.push(`<p>${inlineFormat(trimmed)}</p>`);
  }
  closeLists();
  return out.length > 0 ? out.join('') : '<p></p>';
}

export function htmlToMarkdown(html: string): string {
  let s = html;
  s = s.replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, (_, inner) => `# ${htmlToPlainText(inner)}\n\n`);
  s = s.replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, (_, inner) => `## ${htmlToPlainText(inner)}\n\n`);
  s = s.replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, (_, inner) => `### ${htmlToPlainText(inner)}\n\n`);
  s = s.replace(/<h4[^>]*>([\s\S]*?)<\/h4>/gi, (_, inner) => `#### ${htmlToPlainText(inner)}\n\n`);
  s = s.replace(/<strong[^>]*>([\s\S]*?)<\/strong>/gi, '**$1**');
  s = s.replace(/<b[^>]*>([\s\S]*?)<\/b>/gi, '**$1**');
  s = s.replace(/<em[^>]*>([\s\S]*?)<\/em>/gi, '*$1*');
  s = s.replace(/<i[^>]*>([\s\S]*?)<\/i>/gi, '*$1*');
  s = s.replace(/<s[^>]*>([\s\S]*?)<\/s>/gi, '~~$1~~');
  s = s.replace(/<a[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)');
  s = s.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, (_, inner) => `- ${htmlToPlainText(inner)}\n`);
  s = s.replace(/<blockquote[^>]*>([\s\S]*?)<\/blockquote>/gi, (_, inner) => `> ${htmlToPlainText(inner)}\n\n`);
  s = s.replace(/<hr\s*\/?>/gi, '\n---\n\n');
  s = s.replace(/<div[^>]*data-page-break="true"[^>]*>[\s\S]*?<\/div>/gi, '\n---\n\n');
  s = s.replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<\/p>/gi, '\n\n');
  s = s.replace(/<[^>]+>/g, '');
  return decodeHtmlEntities(s).replace(/\n{3,}/g, '\n\n').trim();
}

export function htmlToRtf(html: string): string {
  const blocks = extractSimpleBlocks(html);
  const escapeRtf = (text: string) =>
    text
      .replace(/\\/g, '\\\\')
      .replace(/\{/g, '\\{')
      .replace(/\}/g, '\\}')
      .replace(/[^\x20-\x7E]/g, (ch) => `\\u${ch.charCodeAt(0)}?`);

  const body = blocks
    .map((b) => {
      if (b.type === 'pagebreak') return '\\page ';
      const prefix =
        b.type === 'h1'
          ? '\\fs36\\b '
          : b.type === 'h2'
            ? '\\fs28\\b '
            : b.type === 'h3'
              ? '\\fs24\\b '
              : '\\fs22 ';
      const suffix = b.type.startsWith('h') ? '\\b0\\fs22\\par\n' : '\\par\n';
      const txt = b.type === 'li' ? `\\bullet  ${escapeRtf(b.text)}` : escapeRtf(b.text);
      return `${prefix}${txt}${suffix}`;
    })
    .join('');

  return `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Calibri;}}\n${body}}`;
}

export function rtfToHtml(rtf: string): string {
  // Strip header groups like fonttbl, colortbl, stylesheet, info
  let cleaned = rtf
    .replace(/\{\\\*\\[^}]*\}/g, '')
    .replace(/\{\\(?:fonttbl|colortbl|stylesheet|info|pict)[^{}]*(?:\{[^{}]*\}[^{}]*)*\}/g, '');

  // Convert paragraph breaks and hex/unicode escapes
  cleaned = cleaned
    .replace(/\\par\b/g, '\n')
    .replace(/\\line\b/g, '\n')
    .replace(/\\tab\b/g, '    ')
    .replace(/\\u(-?\d+)\??/g, (_, code) => {
      const n = Number(code);
      return String.fromCharCode(n < 0 ? n + 65536 : n);
    })
    .replace(/\\'([0-9a-fA-F]{2})/g, (_, hex) => String.fromCharCode(Number.parseInt(hex, 16)))
    .replace(/\\[a-zA-Z]+-?\d* ?/g, '')
    .replace(/[{}]/g, '');

  return plainTextToHtml(cleaned.trim());
}

interface ParsedInlineRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sub?: boolean;
  superScript?: boolean;
  color?: string;
  fontFamily?: string;
  fontSizeHalfPts?: number;
}

interface ParsedBlock {
  type: 'h1' | 'h2' | 'h3' | 'h4' | 'p' | 'li' | 'ol-li' | 'blockquote' | 'table' | 'pagebreak';
  text: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  runs: ParsedInlineRun[];
  rows?: string[][];
  footnoteIds?: string[];
}

function parseFontSizeToHalfPts(raw: string): number | undefined {
  const m = raw.trim().match(/^(\d+(?:\.\d+)?)\s*(pt|px)$/i);
  if (!m) return undefined;
  const num = Number(m[1]);
  if (!Number.isFinite(num) || num <= 0) return undefined;
  const pt = m[2].toLowerCase() === 'px' ? num * 0.75 : num;
  return Math.max(8, Math.round(pt * 2));
}

function parseInlineRuns(innerHtml: string): ParsedInlineRun[] {
  const runs: ParsedInlineRun[] = [];
  const tokenRegex = /(<\/?(?:strong|b|em|i|u|s|del|sub|sup|span|a|br)[^>]*>)/gi;
  const parts = innerHtml.split(tokenRegex);

  const state = {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    sub: false,
    superScript: false,
  };
  const spanStack: Array<{
    color?: string;
    fontFamily?: string;
    fontSizeHalfPts?: number;
  }> = [];

  const currentSpanStyle = () => {
    let color: string | undefined;
    let fontFamily: string | undefined;
    let fontSizeHalfPts: number | undefined;
    for (const frame of spanStack) {
      if (frame.color) color = frame.color;
      if (frame.fontFamily) fontFamily = frame.fontFamily;
      if (frame.fontSizeHalfPts) fontSizeHalfPts = frame.fontSizeHalfPts;
    }
    return {color, fontFamily, fontSizeHalfPts};
  };

  for (const part of parts) {
    if (!part) continue;
    if (part.startsWith('<')) {
      const lower = part.toLowerCase();
      if (lower.startsWith('<br')) {
        runs.push({text: '\n'});
      } else if (lower.startsWith('<strong') || lower.startsWith('<b')) {
        state.bold = true;
      } else if (lower === '</strong>' || lower === '</b>') {
        state.bold = false;
      } else if (lower.startsWith('<em') || lower.startsWith('<i')) {
        state.italic = true;
      } else if (lower === '</em>' || lower === '</i>') {
        state.italic = false;
      } else if (lower.startsWith('<u')) {
        state.underline = true;
      } else if (lower === '</u>') {
        state.underline = false;
      } else if (lower.startsWith('<s') && !lower.startsWith('<sp') && !lower.startsWith('<su') && !lower.startsWith('<st')) {
        state.strike = true;
      } else if (lower === '</s>' || lower === '</del>') {
        state.strike = false;
      } else if (lower.startsWith('<sub')) {
        state.sub = true;
      } else if (lower === '</sub>') {
        state.sub = false;
      } else if (lower.startsWith('<sup')) {
        state.superScript = true;
      } else if (lower === '</sup>') {
        state.superScript = false;
      } else if (lower.startsWith('<span')) {
        const styleAttrRaw =
          part.match(/\bstyle="([^"]*)"/i)?.[1] ?? part.match(/\bstyle='([^']*)'/i)?.[1] ?? '';
        const decodedStyle = decodeHtmlEntities(styleAttrRaw);
        const colorMatch = decodedStyle.match(/(?:^|;)\s*color\s*:\s*#([0-9a-fA-F]{6})/i);
        const dataFontMatch = part.match(/data-doc-font=["']([^"']+)["']/i);
        const cssFontMatch = decodedStyle.match(/(?:^|;)\s*font-family\s*:\s*([^;]+)/i);
        const rawFont = dataFontMatch
          ? decodeHtmlEntities(dataFontMatch[1])
          : cssFontMatch
            ? cssFontMatch[1].trim()
            : null;
        const canonicalFont = canonicalizeWordFontFamily(rawFont) ?? undefined;
        const sizeMatch = decodedStyle.match(/(?:^|;)\s*font-size\s*:\s*([0-9.]+(?:pt|px))/i);
        const fontSizeHalfPts = sizeMatch ? parseFontSizeToHalfPts(sizeMatch[1]) : undefined;

        spanStack.push({
          color: colorMatch ? colorMatch[1].toUpperCase() : undefined,
          fontFamily: canonicalFont,
          fontSizeHalfPts,
        });
      } else if (lower === '</span>') {
        spanStack.pop();
      }
      continue;
    }
    const text = decodeHtmlEntities(part.replace(/<[^>]+>/g, ''));
    if (text) {
      const span = currentSpanStyle();
      runs.push({
        text,
        bold: state.bold,
        italic: state.italic,
        underline: state.underline,
        strike: state.strike,
        sub: state.sub,
        superScript: state.superScript,
        color: span.color,
        fontFamily: span.fontFamily,
        fontSizeHalfPts: span.fontSizeHalfPts,
      });
    }
  }
  return runs;
}

function extractAlign(tagHeader: string): ParsedBlock['align'] {
  const m = tagHeader.match(/text-align:\s*(left|center|right|justify)/i);
  return m ? (m[1].toLowerCase() as ParsedBlock['align']) : undefined;
}

export function extractSimpleBlocks(html: string): ParsedBlock[] {
  const blocks: ParsedBlock[] = [];
  const blockRegex =
    /<div[^>]*data-page-break="true"[^>]*>[\s\S]*?<\/div>|<table[^>]*>[\s\S]*?<\/table>|<(h[1-4]|p|li|blockquote)([^>]*)>([\s\S]*?)<\/\1>/gi;

  let match: RegExpExecArray | null;
  while ((match = blockRegex.exec(html)) !== null) {
    const full = match[0];
    if (full.includes('data-page-break="true"')) {
      blocks.push({type: 'pagebreak', text: '', runs: []});
      continue;
    }
    if (full.toLowerCase().startsWith('<table')) {
      const rows: string[][] = [];
      const rowMatches = full.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) ?? [];
      for (const rm of rowMatches) {
        const cells = (rm.match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) ?? []).map((c) =>
          htmlToPlainText(c),
        );
        if (cells.length > 0) rows.push(cells);
      }
      if (rows.length > 0) {
        const tableFnIds = extractFootnoteIdsFromHtml(full);
        blocks.push({
          type: 'table',
          text: '',
          runs: [],
          rows,
          footnoteIds: tableFnIds.length > 0 ? tableFnIds : undefined,
        });
      }
      continue;
    }

    const tag = match[1].toLowerCase() as ParsedBlock['type'];
    const attrs = match[2] || '';
    const inner = match[3] || '';
    const text = htmlToPlainText(inner);
    const runs = parseInlineRuns(inner);
    const fnIds = extractFootnoteIdsFromHtml(inner);
    blocks.push({
      type: tag,
      text,
      align: extractAlign(attrs),
      runs: runs.length > 0 ? runs : [{text}],
      footnoteIds: fnIds.length > 0 ? fnIds : undefined,
    });
  }

  if (blocks.length === 0) {
    const plain = htmlToPlainText(html);
    if (plain) {
      blocks.push({type: 'p', text: plain, runs: [{text: plain}]});
    }
  }
  return blocks;
}

function extractFootnoteIdsFromHtml(fragment: string): string[] {
  const ids: string[] = [];
  const regex = /data-note-ref="([^"]+)"[^>]*data-note-kind="footnote"|data-note-kind="footnote"[^>]*data-note-ref="([^"]+)"/gi;
  let m: RegExpExecArray | null;
  while ((m = regex.exec(fragment)) !== null) {
    const id = m[1] || m[2];
    if (id && !ids.includes(id)) ids.push(id);
  }
  return ids;
}

export async function htmlToDocxBytes(
  html: string,
  layout: DocPageLayout = DEFAULT_DOC_LAYOUT,
  notes: DocNoteItem[] = [],
  originalDocxBuffer?: ArrayBuffer | Uint8Array | null,
): Promise<Uint8Array> {
  return exportHtmlToDocxBytes(html, layout, notes, originalDocxBuffer);
}

export async function htmlToPdfBytes(
  html: string,
  title: string,
  layout: DocPageLayout = DEFAULT_DOC_LAYOUT,
  notes: DocNoteItem[] = [],
): Promise<Uint8Array> {
  return exportHtmlToPdfHeadless(html, title, layout, notes);
}
