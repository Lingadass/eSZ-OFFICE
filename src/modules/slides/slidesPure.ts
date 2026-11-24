import JSZip from 'jszip';
import {PDFDocument, rgb, StandardFonts} from 'pdf-lib';
import PptxGenJS from 'pptxgenjs';
import {
  fitTextToSlideBox,
  getSlideDimensions,
  inferAspectRatioFromEmu,
  resolveImportedFontName,
  type ImportedThemeFonts,
} from './slidesTextFit';
import {
  SLIDE_HEIGHT,
  SLIDE_INCH_H,
  SLIDE_INCH_W,
  SLIDE_WIDTH,
  type ChartElement,
  type DeckTemplate,
  type ImageElement,
  type PersistedSlidesMeta,
  type PresentationDeck,
  type ShapeElement,
  type ShapeKind,
  type SlideAspectRatio,
  type SlideChartKind,
  type SlideElement,
  type SlideItem,
  type SlideLayoutKind,
  type SlideTheme,
  type TableElement,
  type TextElement,
} from './slidesTypes';

let idCounter = 1;
export function createSlideUid(prefix = 'el'): string {
  idCounter += 1;
  return `${prefix}_${Date.now().toString(36)}_${idCounter.toString(36)}_${Math.random()
    .toString(36)
    .slice(2, 6)}`;
}

export const SLIDE_THEMES: SlideTheme[] = [
  {
    id: 'executive-slate',
    name: 'Executive Slate',
    description: 'Clean crisp paper with deep navy ink and cobalt accents.',
    background: '#FFFFFF',
    surface: '#F1F5F9',
    primary: '#1E3A8A',
    accent: '#D97706',
    headingColor: '#0F172A',
    bodyColor: '#334155',
    mutedColor: '#64748B',
    headingFont: 'Plus Jakarta Sans',
    bodyFont: 'DM Sans',
    chartPalette: ['#1D4ED8', '#0284C7', '#0D9488', '#D97706', '#6366F1'],
  },
  {
    id: 'midnight-keynote',
    name: 'Midnight Keynote',
    description: 'Deep obsidian canvas built for conference projectors.',
    background: '#0F172A',
    surface: '#1E293B',
    primary: '#38BDF8',
    accent: '#F59E0B',
    headingColor: '#F8FAFC',
    bodyColor: '#CBD5E1',
    mutedColor: '#94A3B8',
    headingFont: 'Plus Jakarta Sans',
    bodyFont: 'DM Sans',
    chartPalette: ['#38BDF8', '#34D399', '#FBBF24', '#F472B6', '#A78BFA'],
  },
  {
    id: 'forest-editorial',
    name: 'Forest Editorial',
    description: 'Editorial ivory with deep emerald headings and warm stone.',
    background: '#FAFBF9',
    surface: '#EDF2EE',
    primary: '#065F46',
    accent: '#B45309',
    headingColor: '#064E3B',
    bodyColor: '#374151',
    mutedColor: '#6B7280',
    headingFont: 'Georgia',
    bodyFont: 'DM Sans',
    chartPalette: ['#047857', '#0D9488', '#D97706', '#2563EB', '#64748B'],
  },
  {
    id: 'amber-studio',
    name: 'Amber Studio',
    description: 'Warm alabaster canvas with terracotta-orange presentation highlights.',
    background: '#FFFDF9',
    surface: '#FEF3E2',
    primary: '#C2410C',
    accent: '#1D4ED8',
    headingColor: '#1C1917',
    bodyColor: '#44403C',
    mutedColor: '#78716C',
    headingFont: 'Plus Jakarta Sans',
    bodyFont: 'DM Sans',
    chartPalette: ['#EA580C', '#D97706', '#0284C7', '#059669', '#7C3AED'],
  },
  {
    id: 'ocean-cobalt',
    name: 'Ocean Cobalt',
    description: 'Cool frost surface with rich royal blue visual structure.',
    background: '#F8FAFC',
    surface: '#E0F2FE',
    primary: '#0369A1',
    accent: '#0F766E',
    headingColor: '#0C4A6E',
    bodyColor: '#334155',
    mutedColor: '#64748B',
    headingFont: 'Trebuchet MS',
    bodyFont: 'DM Sans',
    chartPalette: ['#0284C7', '#0EA5E9', '#0D9488', '#4F46E5', '#F59E0B'],
  },
  {
    id: 'nordic-charcoal',
    name: 'Nordic Charcoal',
    description: 'Minimal monochrome graphite with crimson focal highlights.',
    background: '#18181B',
    surface: '#27272A',
    primary: '#E11D48',
    accent: '#38BDF8',
    headingColor: '#FAFAFA',
    bodyColor: '#D4D4D8',
    mutedColor: '#A1A1AA',
    headingFont: 'Plus Jakarta Sans',
    bodyFont: 'DM Sans',
    chartPalette: ['#E11D48', '#F97316', '#38BDF8', '#22C55E', '#A855F7'],
  },
];

export function getSlideTheme(themeId: string): SlideTheme {
  return SLIDE_THEMES.find((t) => t.id === themeId) ?? SLIDE_THEMES[0];
}

export function formatListDisplayText(
  text: string,
  listStyle?: 'none' | 'bullet' | 'numbered',
): string {
  if (!listStyle || listStyle === 'none') return text;
  return text
    .split('\n')
    .map((line, idx) => {
      const trimmed = line.trim();
      if (!trimmed) return '';
      const stripped = trimmed.replace(/^([•\-*]|\d+\.)\s+/, '');
      return listStyle === 'bullet' ? `•  ${stripped}` : `${idx + 1}.  ${stripped}`;
    })
    .join('\n');
}

export function createSlideFromLayout(
  layout: SlideLayoutKind,
  theme: SlideTheme,
  overrides?: {title?: string; subtitle?: string; body?: string; notes?: string},
): SlideItem {
  const slideId = createSlideUid('slide');
  const elements: SlideElement[] = [];

  if (layout === 'title') {
    elements.push(
      {
        id: createSlideUid('shp'),
        type: 'shape',
        shapeKind: 'rect',
        x: 72,
        y: 138,
        width: 84,
        height: 6,
        rotation: 0,
        fill: theme.primary,
        stroke: 'transparent',
        strokeWidth: 0,
      },
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'title',
        x: 72,
        y: 162,
        width: 816,
        height: 112,
        rotation: 0,
        text: overrides?.title ?? 'Presentation Title',
        fontFamily: theme.headingFont,
        fontSize: 44,
        bold: true,
        align: 'left',
        color: theme.headingColor,
        lineHeight: 1.15,
      },
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'subtitle',
        x: 72,
        y: 286,
        width: 760,
        height: 76,
        rotation: 0,
        text: overrides?.subtitle ?? 'Subtitle, presenter name, or date',
        fontFamily: theme.bodyFont,
        fontSize: 20,
        align: 'left',
        color: theme.mutedColor,
        lineHeight: 1.4,
      },
    );
  } else if (layout === 'title-content') {
    elements.push(
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'title',
        x: 56,
        y: 42,
        width: 848,
        height: 62,
        rotation: 0,
        text: overrides?.title ?? 'Slide Heading',
        fontFamily: theme.headingFont,
        fontSize: 32,
        bold: true,
        align: 'left',
        color: theme.headingColor,
        lineHeight: 1.2,
      },
      {
        id: createSlideUid('shp'),
        type: 'shape',
        shapeKind: 'rect',
        x: 56,
        y: 108,
        width: 848,
        height: 2,
        rotation: 0,
        fill: theme.surface,
        stroke: 'transparent',
        strokeWidth: 0,
      },
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'body',
        x: 56,
        y: 132,
        width: 848,
        height: 336,
        rotation: 0,
        text:
          overrides?.body ??
          'Key takeaway or supporting metric\nSecond concise point for the audience\nThird actionable next step',
        fontFamily: theme.bodyFont,
        fontSize: 20,
        align: 'left',
        color: theme.bodyColor,
        lineHeight: 1.55,
        listStyle: 'bullet',
      },
    );
  } else if (layout === 'two-column') {
    elements.push(
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'title',
        x: 56,
        y: 40,
        width: 848,
        height: 58,
        rotation: 0,
        text: overrides?.title ?? 'Two-Column Overview',
        fontFamily: theme.headingFont,
        fontSize: 32,
        bold: true,
        align: 'left',
        color: theme.headingColor,
      },
      {
        id: createSlideUid('shp'),
        type: 'shape',
        shapeKind: 'rounded-rect',
        x: 56,
        y: 120,
        width: 408,
        height: 356,
        rotation: 0,
        fill: theme.surface,
        stroke: 'transparent',
        strokeWidth: 0,
        cornerRadius: 12,
      },
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'body',
        x: 76,
        y: 140,
        width: 368,
        height: 316,
        rotation: 0,
        text: overrides?.body ?? 'Primary focus area\nCore milestones & deliverables\nResource allocation',
        fontFamily: theme.bodyFont,
        fontSize: 18,
        color: theme.bodyColor,
        lineHeight: 1.55,
        listStyle: 'bullet',
      },
      {
        id: createSlideUid('shp'),
        type: 'shape',
        shapeKind: 'rounded-rect',
        x: 496,
        y: 120,
        width: 408,
        height: 356,
        rotation: 0,
        fill: theme.surface,
        stroke: 'transparent',
        strokeWidth: 0,
        cornerRadius: 12,
      },
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'body',
        x: 516,
        y: 140,
        width: 368,
        height: 316,
        rotation: 0,
        text:
          overrides?.subtitle ??
          'Secondary column details\nRisks & mitigation plan\nTarget launch window',
        fontFamily: theme.bodyFont,
        fontSize: 18,
        color: theme.bodyColor,
        lineHeight: 1.55,
        listStyle: 'bullet',
      },
    );
  } else if (layout === 'section') {
    elements.push(
      {
        id: createSlideUid('shp'),
        type: 'shape',
        shapeKind: 'rect',
        x: 0,
        y: 0,
        width: 320,
        height: SLIDE_HEIGHT,
        rotation: 0,
        fill: theme.primary,
        stroke: 'transparent',
        strokeWidth: 0,
      },
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'subtitle',
        x: 368,
        y: 182,
        width: 520,
        height: 36,
        rotation: 0,
        text: overrides?.subtitle ?? 'SECTION BREAK',
        fontFamily: theme.bodyFont,
        fontSize: 14,
        bold: true,
        color: theme.primary,
      },
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'title',
        x: 368,
        y: 222,
        width: 530,
        height: 120,
        rotation: 0,
        text: overrides?.title ?? 'Section Headline',
        fontFamily: theme.headingFont,
        fontSize: 40,
        bold: true,
        color: theme.headingColor,
        lineHeight: 1.15,
      },
    );
  } else if (layout === 'comparison') {
    elements.push(
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'title',
        x: 56,
        y: 38,
        width: 848,
        height: 54,
        rotation: 0,
        text: overrides?.title ?? 'Comparison Matrix',
        fontFamily: theme.headingFont,
        fontSize: 30,
        bold: true,
        color: theme.headingColor,
      },
      {
        id: createSlideUid('tbl'),
        type: 'table',
        x: 56,
        y: 120,
        width: 848,
        height: 330,
        rotation: 0,
        rows: [
          ['Capability', 'Current State', 'Target Release'],
          ['Offline Processing', 'Local Storage', 'Full PWA + OPFS'],
          ['Export Formats', 'Standard', 'PPTX, PDF & PNG'],
          ['Privacy Model', 'Zero Telemetry', '100% On-Device'],
        ],
        headerBg: theme.primary,
        headerColor: '#FFFFFF',
        cellBg: theme.surface,
        cellColor: theme.bodyColor,
        borderColor: '#CBD5E1',
        fontSize: 15,
      },
    );
  } else if (layout === 'image-caption') {
    elements.push(
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'title',
        x: 56,
        y: 40,
        width: 848,
        height: 54,
        rotation: 0,
        text: overrides?.title ?? 'Metrics & Visual Spotlight',
        fontFamily: theme.headingFont,
        fontSize: 30,
        bold: true,
        color: theme.headingColor,
      },
      {
        id: createSlideUid('cht'),
        type: 'chart',
        chartKind: 'bar',
        x: 56,
        y: 118,
        width: 520,
        height: 356,
        rotation: 0,
        title: 'Quarterly Adoption',
        labels: ['Q1', 'Q2', 'Q3', 'Q4'],
        values: [42, 68, 84, 112],
        colors: theme.chartPalette,
      },
      {
        id: createSlideUid('txt'),
        type: 'text',
        role: 'body',
        x: 604,
        y: 136,
        width: 300,
        height: 320,
        rotation: 0,
        text:
          overrides?.body ??
          'Steady quarter-over-quarter momentum\nOffline workflows account for 100% of sessions\nZero cloud dependency reduces latency',
        fontFamily: theme.bodyFont,
        fontSize: 18,
        color: theme.bodyColor,
        lineHeight: 1.55,
        listStyle: 'bullet',
      },
    );
  }

  return {
    id: slideId,
    layout,
    background: theme.background,
    transition: 'fade',
    notes: overrides?.notes ?? '',
    elements,
  };
}

export function createBlankDeck(title = 'Untitled Presentation', themeId = 'executive-slate'): PresentationDeck {
  const theme = getSlideTheme(themeId);
  const firstSlide = createSlideFromLayout('title', theme, {
    title: title.replace(/\.pptx$/i, ''),
    subtitle: 'Created offline with eSZ OFFICE',
    notes: 'Introduce the deck topic and core objectives.',
  });
  return {
    title: title.replace(/\.pptx$/i, ''),
    themeId: theme.id,
    aspectRatio: '16:9',
    master: {
      footerText: 'eSZ OFFICE · Offline Presentation',
      showSlideNumbers: true,
      showAccentBar: true,
    },
    slides: [firstSlide],
    activeSlideId: firstSlide.id,
  };
}

export const DECK_TEMPLATES: DeckTemplate[] = [
  {
    id: 'product-roadmap',
    name: 'Product Roadmap',
    category: 'Strategy',
    description: '4-slide executive roadmap with milestones, metrics chart, and comparison table.',
    themeId: 'executive-slate',
    buildSlides: (theme) => [
      createSlideFromLayout('title', theme, {
        title: 'Product Roadmap & Vision',
        subtitle: 'Offline-first productivity suite · Quarterly Strategy',
        notes: 'Welcome stakeholders and walk through our quarterly milestones.',
      }),
      createSlideFromLayout('two-column', theme, {
        title: 'Core Architecture Pillars',
        body: 'Zero-upload on-device vault (IndexedDB + OPFS)\nModular PDF, Docs, Sheets & Slides engines\nService Worker precaching for offline reliability',
        subtitle: 'Strict Content Security Policy\nOpen-source permissive licences only\nResponsive touch & desktop keyboard workflows',
        notes: 'Highlight privacy-by-design and offline resilience.',
      }),
      createSlideFromLayout('image-caption', theme, {
        title: 'Adoption & Performance Trajectory',
        body: 'Faster cold loads with lazy-loaded modules\n100% local rendering keeps latency below 50ms\nFull export fidelity across DOCX, XLSX, PPTX & PDF',
        notes: 'Point out the steady upward trajectory across Q1-Q4.',
      }),
      createSlideFromLayout('comparison', theme, {
        title: 'Release Readiness Matrix',
        notes: 'Summarize remaining launch gates and next actions.',
      }),
    ],
  },
  {
    id: 'startup-pitch',
    name: 'Executive Pitch Deck',
    category: 'Business',
    description: 'High-contrast dark keynote deck for investor or leadership briefings.',
    themeId: 'midnight-keynote',
    buildSlides: (theme) => [
      createSlideFromLayout('title', theme, {
        title: 'Next-Gen Offline Workspace',
        subtitle: 'Executive Pitch · Confidential Briefing',
        notes: 'Open with the problem statement and market shift.',
      }),
      createSlideFromLayout('section', theme, {
        title: 'The Privacy & Speed Opportunity',
        subtitle: '01 · MARKET CONTEXT',
        notes: 'Pause before diving into customer pain points.',
      }),
      createSlideFromLayout('title-content', theme, {
        title: 'Why Teams Switch to Local-First Tools',
        body: 'Zero vendor lock-in or per-seat subscription fatigue\nWorks seamlessly on flights, field sites, and secure networks\nInstant startup with native open file standards',
        notes: 'Emphasize total cost of ownership and security compliance.',
      }),
      createSlideFromLayout('image-caption', theme, {
        title: 'Growth & Retention Highlights',
        body: '4x organic growth from word-of-mouth referrals\nZero server hosting overhead scales margins cleanly\nCross-platform PWA reaches desktop and Android TWA',
        notes: 'Walk through the Q1–Q4 adoption chart.',
      }),
    ],
  },
  {
    id: 'quarterly-review',
    name: 'Quarterly Business Review',
    category: 'Operations',
    description: 'Editorial green deck structured for quarterly KPI reviews and action items.',
    themeId: 'forest-editorial',
    buildSlides: (theme) => [
      createSlideFromLayout('title', theme, {
        title: 'Quarterly Business Review',
        subtitle: 'Operations, Delivery & Next Quarter Priorities',
        notes: 'Set the agenda for the 30-minute QBR.',
      }),
      createSlideFromLayout('comparison', theme, {
        title: 'Key Deliverable Scorecard',
        notes: 'Review target vs actual completion across workstreams.',
      }),
      createSlideFromLayout('title-content', theme, {
        title: 'Top Priorities for Next Quarter',
        body: 'Complete Play Store TWA packaging & QA checks\nExpand keyboard shortcuts and template library\nMaintain Lighthouse accessibility score above 95',
        notes: 'Assign owners and dates for each priority.',
      }),
    ],
  },
  {
    id: 'workshop-training',
    name: 'Workshop & Training',
    category: 'Education',
    description: 'Warm studio deck built for onboarding sessions and team tutorials.',
    themeId: 'amber-studio',
    buildSlides: (theme) => [
      createSlideFromLayout('title', theme, {
        title: 'Team Onboarding Workshop',
        subtitle: 'Hands-On Guide · Best Practices & Exercises',
        notes: 'Introduce facilitators and confirm everyone has sample files ready.',
      }),
      createSlideFromLayout('two-column', theme, {
        title: 'Session Agenda',
        body: 'Part 1: Local File Manager & Backups\nPart 2: PDF Annotation & Signing\nPart 3: Rich Documents & Spreadsheets',
        subtitle: 'Part 4: Slide Design & Charts\nPart 5: Presenter Mode & Export\nQ&A + Practical Wrap-Up',
        notes: 'Outline timing for each hands-on module.',
      }),
      createSlideFromLayout('title-content', theme, {
        title: 'Quick Keyboard Shortcuts',
        body: 'Ctrl+S saves your deck immediately to device storage\nCtrl+Z / Ctrl+Y undoes and redoes canvas edits\nF5 launches full-screen Presenter Mode with speaker notes',
        notes: 'Encourage participants to try F5 before finishing.',
      }),
    ],
  },
];

export function createDeckFromTemplate(templateId: string, customTitle?: string): PresentationDeck {
  const tpl = DECK_TEMPLATES.find((t) => t.id === templateId) ?? DECK_TEMPLATES[0];
  const theme = getSlideTheme(tpl.themeId);
  const slides = tpl.buildSlides(theme);
  return {
    title: customTitle || tpl.name,
    themeId: theme.id,
    aspectRatio: '16:9',
    master: {
      footerText: 'eSZ OFFICE',
      showSlideNumbers: true,
      showAccentBar: true,
    },
    slides,
    activeSlideId: slides[0].id,
  };
}

/**
 * Parses Markdown / structured outline text into ready-to-present slides.
 * Supports:
 * - `# Slide Title` / `## Subtitle` / `- Bullet` / `> Speaker notes`
 * - `Slide 1: Title` lines (used by starter seeds)
 */
export function parseOutlineToSlides(outlineText: string, themeId = 'executive-slate'): SlideItem[] {
  const theme = getSlideTheme(themeId);
  const lines = outlineText
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  if (lines.length === 0) {
    return [createSlideFromLayout('title', theme)];
  }

  interface DraftSlide {
    title: string;
    subtitle: string;
    bullets: string[];
    notes: string[];
  }

  const drafts: DraftSlide[] = [];
  let current: DraftSlide | null = null;

  for (const line of lines) {
    const slidePrefix = line.match(/^slide\s+\d+\s*[:\-–]\s*(.+)$/i);
    if (slidePrefix) {
      if (current) drafts.push(current);
      current = {title: slidePrefix[1].trim(), subtitle: '', bullets: [], notes: []};
      continue;
    }

    if (line.startsWith('# ')) {
      if (current) drafts.push(current);
      current = {title: line.slice(2).trim(), subtitle: '', bullets: [], notes: []};
      continue;
    }

    if (!current) {
      current = {title: line.replace(/^#+\s*/, ''), subtitle: '', bullets: [], notes: []};
      continue;
    }

    if (line.startsWith('## ')) {
      current.subtitle = current.subtitle
        ? `${current.subtitle}\n${line.slice(3).trim()}`
        : line.slice(3).trim();
    } else if (line.startsWith('> ')) {
      current.notes.push(line.slice(2).trim());
    } else if (/^([-*•]|\d+\.)\s+/.test(line)) {
      current.bullets.push(line.replace(/^([-*•]|\d+\.)\s+/, '').trim());
    } else if (current.bullets.length === 0 && !current.subtitle) {
      current.subtitle = line;
    } else {
      current.bullets.push(line);
    }
  }
  if (current) drafts.push(current);

  return drafts.map((d, idx) => {
    const notes = d.notes.join('\n');
    if (idx === 0 && d.bullets.length === 0) {
      return createSlideFromLayout('title', theme, {
        title: d.title,
        subtitle: d.subtitle || 'Generated from outline · eSZ OFFICE',
        notes,
      });
    }
    if (d.bullets.length >= 6) {
      const half = Math.ceil(d.bullets.length / 2);
      return createSlideFromLayout('two-column', theme, {
        title: d.title,
        body: d.bullets.slice(0, half).join('\n'),
        subtitle: d.bullets.slice(half).join('\n'),
        notes,
      });
    }
    if (d.bullets.length === 0 && d.subtitle) {
      return createSlideFromLayout('section', theme, {
        title: d.title,
        subtitle: d.subtitle,
        notes,
      });
    }
    return createSlideFromLayout('title-content', theme, {
      title: d.title,
      body:
        d.bullets.length > 0
          ? d.bullets.join('\n')
          : d.subtitle || 'Add supporting bullet points or visual elements here',
      notes,
    });
  });
}

export function applyThemeToDeck(deck: PresentationDeck, nextThemeId: string): PresentationDeck {
  const prevTheme = getSlideTheme(deck.themeId);
  const nextTheme = getSlideTheme(nextThemeId);
  return {
    ...deck,
    themeId: nextTheme.id,
    slides: deck.slides.map((slide) => ({
      ...slide,
      background:
        slide.background.toLowerCase() === prevTheme.background.toLowerCase()
          ? nextTheme.background
          : nextTheme.background,
      elements: slide.elements.map((el) => {
        if (el.type === 'text') {
          const nextColor =
            el.role === 'title'
              ? nextTheme.headingColor
              : el.role === 'subtitle'
                ? nextTheme.mutedColor
                : nextTheme.bodyColor;
          return {
            ...el,
            color: nextColor,
            fontFamily: el.role === 'title' ? nextTheme.headingFont : nextTheme.bodyFont,
          };
        }
        if (el.type === 'shape') {
          const isPrimary = el.fill.toLowerCase() === prevTheme.primary.toLowerCase();
          const isSurface = el.fill.toLowerCase() === prevTheme.surface.toLowerCase();
          return {
            ...el,
            fill: isPrimary ? nextTheme.primary : isSurface ? nextTheme.surface : el.fill,
          };
        }
        if (el.type === 'table') {
          return {
            ...el,
            headerBg: nextTheme.primary,
            cellBg: nextTheme.surface,
            cellColor: nextTheme.bodyColor,
          };
        }
        if (el.type === 'chart') {
          return {
            ...el,
            colors: nextTheme.chartPalette,
          };
        }
        return el;
      }),
    })),
  };
}

export function reorderSlides(slides: SlideItem[], fromIndex: number, toIndex: number): SlideItem[] {
  if (
    fromIndex < 0 ||
    toIndex < 0 ||
    fromIndex >= slides.length ||
    toIndex >= slides.length ||
    fromIndex === toIndex
  ) {
    return slides;
  }
  const copy = [...slides];
  const [moved] = copy.splice(fromIndex, 1);
  copy.splice(toIndex, 0, moved);
  return copy;
}

export function duplicateSlideItem(slide: SlideItem): SlideItem {
  return {
    ...slide,
    id: createSlideUid('slide'),
    elements: slide.elements.map((el) => ({
      ...el,
      id: createSlideUid(el.type.slice(0, 3)),
      ...(el.type === 'table' ? {rows: el.rows.map((r) => [...r])} : {}),
      ...(el.type === 'chart'
        ? {labels: [...el.labels], values: [...el.values], colors: [...el.colors]}
        : {}),
    })),
  };
}

export function alignElementOnSlide(
  el: SlideElement,
  alignment: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom',
  aspectRatio: SlideAspectRatio = '16:9',
): SlideElement {
  const dims = getSlideDimensions(aspectRatio);
  const margin = 48;
  switch (alignment) {
    case 'left':
      return {...el, x: margin};
    case 'center':
      return {...el, x: Math.round((dims.width - el.width) / 2)};
    case 'right':
      return {...el, x: Math.max(0, dims.width - el.width - margin)};
    case 'top':
      return {...el, y: margin};
    case 'middle':
      return {...el, y: Math.round((dims.height - el.height) / 2)};
    case 'bottom':
      return {...el, y: Math.max(0, dims.height - el.height - margin)};
  }
}

export function moveElementLayer(
  elements: SlideElement[],
  elementId: string,
  direction: 'front' | 'back' | 'forward' | 'backward',
): SlideElement[] {
  const idx = elements.findIndex((e) => e.id === elementId);
  if (idx < 0) return elements;
  const copy = [...elements];
  const [item] = copy.splice(idx, 1);
  if (direction === 'front') {
    copy.push(item);
  } else if (direction === 'back') {
    copy.unshift(item);
  } else if (direction === 'forward') {
    copy.splice(Math.min(copy.length, idx + 1), 0, item);
  } else {
    copy.splice(Math.max(0, idx - 1), 0, item);
  }
  return copy;
}

function stripHexHash(hex: string, fallback = '0F172A'): string {
  const clean = (hex || '').replace(/^#/, '').trim();
  if (/^[0-9a-fA-F]{6}$/.test(clean)) return clean.toUpperCase();
  if (/^[0-9a-fA-F]{3}$/.test(clean)) {
    return clean
      .split('')
      .map((c) => c + c)
      .join('')
      .toUpperCase();
  }
  return fallback;
}

function pxToIn(px: number, slidePx: number, slideIn: number): number {
  return Number(((px / slidePx) * slideIn).toFixed(3));
}

/**
 * Exports a PresentationDeck to a genuine OpenXML .pptx binary using PptxGenJS (MIT).
 */
export async function deckToPptxBytes(deck: PresentationDeck): Promise<Uint8Array> {
  const Ctor =
    (PptxGenJS as unknown as {default?: typeof PptxGenJS}).default ?? PptxGenJS;
  const pptx = new Ctor();
  const dims = getSlideDimensions(deck.aspectRatio);
  pptx.layout = dims.aspectRatio === '4:3' ? 'LAYOUT_4x3' : 'LAYOUT_16x9';
  pptx.title = deck.title || 'Presentation';
  pptx.author = 'eSZ OFFICE';
  pptx.subject = 'Created offline with eSZ OFFICE';

  const theme = getSlideTheme(deck.themeId);

  deck.slides.forEach((slideItem, slideIdx) => {
    const slide = pptx.addSlide();
    slide.background = {color: stripHexHash(slideItem.background, 'FFFFFF')};

    if (deck.master.showAccentBar) {
      slide.addShape('rect', {
        x: 0,
        y: 0,
        w: dims.inchW,
        h: 0.06,
        fill: {color: stripHexHash(theme.primary, '1E3A8A')},
        line: {type: 'none'},
      });
    }

    for (const el of slideItem.elements) {
      const x = pxToIn(el.x, dims.width, dims.inchW);
      const y = pxToIn(el.y, dims.height, dims.inchH);
      const w = Math.max(0.15, pxToIn(el.width, dims.width, dims.inchW));
      const h = Math.max(0.15, pxToIn(el.height, dims.height, dims.inchH));

      if (el.type === 'text') {
        const displayText = formatListDisplayText(el.text, el.listStyle);
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
        slide.addText(displayText, {
          x,
          y,
          w,
          h,
          fontFace: el.fontFamily || 'Arial',
          fontSize: Math.max(8, Math.round(fit.fittedFontSize * 0.75)),
          bold: Boolean(el.bold),
          italic: Boolean(el.italic),
          underline: el.underline ? {style: 'sng'} : undefined,
          align: el.align ?? 'left',
          color: stripHexHash(el.color, '0F172A'),
          fill:
            el.fill && el.fill !== 'transparent'
              ? {color: stripHexHash(el.fill, 'FFFFFF')}
              : undefined,
          rotate: el.rotation || 0,
          valign: 'top',
          margin: 4,
          fit: 'shrink',
        });
      } else if (el.type === 'shape') {
        const shapeMap: Record<ShapeKind, PptxGenJS.SHAPE_NAME> = {
          rect: 'rect',
          'rounded-rect': 'roundRect',
          ellipse: 'ellipse',
          triangle: 'triangle',
          'arrow-right': 'rightArrow',
          diamond: 'diamond',
          line: 'line',
        };
        const pptShape = shapeMap[el.shapeKind] ?? 'rect';
        const hasFill = el.shapeKind !== 'line' && el.fill && el.fill !== 'transparent';
        const hasStroke = el.stroke && el.stroke !== 'transparent' && el.strokeWidth > 0;
        if (el.label && el.label.trim().length > 0 && el.shapeKind !== 'line') {
          slide.addText(el.label, {
            shape: pptShape,
            x,
            y,
            w,
            h,
            rotate: el.rotation || 0,
            fill: hasFill ? {color: stripHexHash(el.fill, '1D4ED8')} : undefined,
            line: hasStroke
              ? {color: stripHexHash(el.stroke, '0F172A'), pt: Math.max(1, el.strokeWidth)}
              : {type: 'none'},
            color: stripHexHash(el.labelColor || '#FFFFFF', 'FFFFFF'),
            align: 'center',
            valign: 'middle',
            fontSize: 13,
            bold: true,
          });
        } else {
          slide.addShape(pptShape, {
            x,
            y,
            w,
            h,
            rotate: el.rotation || 0,
            fill: hasFill ? {color: stripHexHash(el.fill, '1D4ED8')} : undefined,
            line: hasStroke
              ? {color: stripHexHash(el.stroke, '0F172A'), pt: Math.max(1, el.strokeWidth)}
              : el.shapeKind === 'line'
                ? {color: stripHexHash(el.fill || theme.primary, '1D4ED8'), pt: 2}
                : {type: 'none'},
          });
        }
      } else if (el.type === 'image') {
        if (el.src.startsWith('data:image/')) {
          slide.addImage({
            data: el.src,
            x,
            y,
            w,
            h,
            rotate: el.rotation || 0,
          });
        }
      } else if (el.type === 'table') {
        const rows = el.rows.map((row, rIdx) =>
          row.map((cell) => ({
            text: cell,
            options: {
              fill: {
                color: stripHexHash(rIdx === 0 ? el.headerBg : el.cellBg, 'F8FAFC'),
              },
              color: stripHexHash(rIdx === 0 ? el.headerColor : el.cellColor, '0F172A'),
              bold: rIdx === 0,
              fontSize: Math.max(9, Math.round(el.fontSize * 0.75)),
              border: {pt: 0.75, color: stripHexHash(el.borderColor, 'CBD5E1')},
              valign: 'middle' as const,
            },
          })),
        );
        if (rows.length > 0) {
          slide.addTable(rows, {
            x,
            y,
            w,
            h,
            autoPage: false,
          });
        }
      } else if (el.type === 'chart') {
        const chartType =
          el.chartKind === 'line'
            ? pptx.ChartType.line
            : el.chartKind === 'pie'
              ? pptx.ChartType.pie
              : pptx.ChartType.bar;
        slide.addChart(
          chartType,
          [
            {
              name: el.title || 'Series 1',
              labels: el.labels.length > 0 ? el.labels : ['A', 'B', 'C'],
              values: el.values.length > 0 ? el.values : [10, 20, 30],
            },
          ],
          {
            x,
            y,
            w,
            h,
            showTitle: Boolean(el.title),
            title: el.title,
            chartColors: (el.colors.length > 0 ? el.colors : theme.chartPalette).map((c) =>
              stripHexHash(c, '1D4ED8'),
            ),
            showLegend: el.chartKind === 'pie',
          },
        );
      }
    }

    // Master footer & slide number
    if (deck.master.footerText.trim()) {
      slide.addText(deck.master.footerText, {
        x: 0.5,
        y: dims.inchH - 0.36,
        w: 7,
        h: 0.25,
        fontSize: 9,
        color: stripHexHash(theme.mutedColor, '64748B'),
      });
    }
    if (deck.master.showSlideNumbers) {
      slide.addText(`${slideIdx + 1}`, {
        x: dims.inchW - 1.0,
        y: dims.inchH - 0.36,
        w: 0.5,
        h: 0.25,
        fontSize: 9,
        align: 'right',
        color: stripHexHash(theme.mutedColor, '64748B'),
      });
    }

    if (slideItem.notes && slideItem.notes.trim()) {
      slide.addNotes(slideItem.notes.trim());
    }
  });

  const output = await pptx.write({outputType: 'uint8array'});
  if (output instanceof Uint8Array) return output;
  if (output instanceof ArrayBuffer) return new Uint8Array(output);
  if (typeof Blob !== 'undefined' && output instanceof Blob) {
    return new Uint8Array(await output.arrayBuffer());
  }
  throw new Error('Unexpected PptxGenJS output type');
}

function decodeXmlEntities(str: string): string {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, num) => String.fromCharCode(Number(num)));
}

function extractParagraphText(xmlChunk: string): string {
  const paraRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
  const paragraphs: string[] = [];
  let pMatch: RegExpExecArray | null;
  while ((pMatch = paraRegex.exec(xmlChunk)) !== null) {
    const pBody = pMatch[1];
    const tRegex = /<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g;
    let line = '';
    let tMatch: RegExpExecArray | null;
    while ((tMatch = tRegex.exec(pBody)) !== null) {
      line += decodeXmlEntities(tMatch[1]);
    }
    paragraphs.push(line);
  }
  return paragraphs.join('\n').trim();
}

function parseEmuBox(
  xmlChunk: string,
  slideCxEmu: number,
  slideCyEmu: number,
  targetW: number,
  targetH: number,
): {x: number; y: number; width: number; height: number} | null {
  const offMatch = xmlChunk.match(/<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/);
  const extMatch = xmlChunk.match(/<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  if (!offMatch || !extMatch) return null;
  const emuPerPxX = slideCxEmu / targetW;
  const emuPerPxY = slideCyEmu / targetH;
  const rawX = Math.round(Number(offMatch[1]) / emuPerPxX);
  const rawY = Math.round(Number(offMatch[2]) / emuPerPxY);
  const x = Math.max(0, Math.min(targetW - 24, rawX));
  const y = Math.max(0, Math.min(targetH - 18, rawY));
  const rawW = Math.max(24, Math.round(Number(extMatch[1]) / emuPerPxX));
  const rawH = Math.max(18, Math.round(Number(extMatch[2]) / emuPerPxY));
  return {
    x,
    y,
    width: Math.min(rawW, targetW - x),
    height: Math.min(rawH, targetH - y),
  };
}

function uint8ToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

/**
 * Reads an OpenXML .pptx binary using JSZip into an editable PresentationDeck.
 * Preserves 16:9 / 4:3 slide dimensions, resolves OOXML theme & run fonts with graceful
 * fallbacks, and auto-shrinks imported text boxes so text never clips or overflows.
 */
export async function pptxBytesToDeck(
  bytes: Uint8Array,
  fileName = 'Presentation.pptx',
): Promise<PresentationDeck> {
  const cleanTitle = fileName.replace(/\.(pptx|ppt|txt|md)$/i, '') || 'Presentation';
  const isZip = bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;

  if (!isZip) {
    const text = new TextDecoder('utf-8', {fatal: false}).decode(bytes).trim();
    if (!text) return createBlankDeck(cleanTitle);
    const slides = parseOutlineToSlides(text, 'executive-slate');
    return {
      title: cleanTitle,
      themeId: 'executive-slate',
      aspectRatio: '16:9',
      master: {
        footerText: 'eSZ OFFICE',
        showSlideNumbers: true,
        showAccentBar: true,
      },
      slides,
      activeSlideId: slides[0].id,
    };
  }

  const zip = await JSZip.loadAsync(bytes);
  const slideEntries = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => {
      const na = Number(a.match(/slide(\d+)\.xml/i)?.[1] ?? 0);
      const nb = Number(b.match(/slide(\d+)\.xml/i)?.[1] ?? 0);
      return na - nb;
    });

  if (slideEntries.length === 0) {
    return createBlankDeck(cleanTitle);
  }

  // Parse ppt/presentation.xml for exact slide EMU size & 16:9 vs 4:3 aspect ratio
  let slideCxEmu = SLIDE_INCH_W * 914400;
  let slideCyEmu = SLIDE_INCH_H * 914400;
  let aspectRatio: SlideAspectRatio = '16:9';

  if (zip.files['ppt/presentation.xml']) {
    const presXml = await zip.files['ppt/presentation.xml'].async('string');
    const sldSzMatch = presXml.match(/<p:sldSz\b([^>]*)\/?>/);
    if (sldSzMatch) {
      const attrs = sldSzMatch[1];
      const cxMatch = attrs.match(/\bcx="(\d+)"/);
      const cyMatch = attrs.match(/\bcy="(\d+)"/);
      const typeMatch = attrs.match(/\btype="([^"]+)"/);
      const parsedCx = cxMatch ? Number(cxMatch[1]) : 0;
      const parsedCy = cyMatch ? Number(cyMatch[1]) : 0;
      aspectRatio = inferAspectRatioFromEmu(parsedCx, parsedCy, typeMatch?.[1]);
      if (parsedCx > 0 && parsedCy > 0) {
        slideCxEmu = parsedCx;
        slideCyEmu = parsedCy;
      } else {
        const inferredDims = getSlideDimensions(aspectRatio);
        slideCxEmu = inferredDims.inchW * 914400;
        slideCyEmu = inferredDims.inchH * 914400;
      }
    }
  }

  const dims = getSlideDimensions(aspectRatio);

  // Parse ppt/theme/theme1.xml for major/minor theme fonts (used for +mj-lt / +mn-lt tokens)
  const themeFonts: ImportedThemeFonts = {};
  const themeEntry = Object.keys(zip.files).find((name) =>
    /^ppt\/theme\/theme\d+\.xml$/i.test(name),
  );
  if (themeEntry) {
    const themeXml = await zip.files[themeEntry].async('string');
    const majorMatch = themeXml.match(
      /<a:majorFont>[\s\S]*?<a:latin\b[^>]*typeface="([^"]+)"/,
    );
    const minorMatch = themeXml.match(
      /<a:minorFont>[\s\S]*?<a:latin\b[^>]*typeface="([^"]+)"/,
    );
    if (majorMatch?.[1]) themeFonts.major = majorMatch[1];
    if (minorMatch?.[1]) themeFonts.minor = minorMatch[1];
  }

  const theme = getSlideTheme('executive-slate');
  const parsedSlides: SlideItem[] = [];

  for (const slidePath of slideEntries) {
    const slideNum = Number(slidePath.match(/slide(\d+)\.xml/i)?.[1] ?? 1);
    const xml = await zip.files[slidePath].async('string');

    // Background color
    const bgMatch = xml.match(/<p:bg>[\s\S]*?<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
    const background = bgMatch ? `#${bgMatch[1]}` : theme.background;

    // Parse slide relationships for embedded media & notes
    const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    const relMap: Record<string, string> = {};
    if (zip.files[relsPath]) {
      const relsXml = await zip.files[relsPath].async('string');
      const relRegex = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
      let rm: RegExpExecArray | null;
      while ((rm = relRegex.exec(relsXml)) !== null) {
        relMap[rm[1]] = rm[2];
      }
    }

    const elements: SlideElement[] = [];

    // 1. Shapes & Text boxes (<p:sp>)
    const spRegex = /<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/g;
    let spMatch: RegExpExecArray | null;
    let textIdx = 0;
    while ((spMatch = spRegex.exec(xml)) !== null) {
      const spXml = spMatch[1];
      const box = parseEmuBox(spXml, slideCxEmu, slideCyEmu, dims.width, dims.height) ?? {
        x: 64,
        y: Math.min(dims.height - 96, 64 + textIdx * 96),
        width: dims.width - 128,
        height: 84,
      };
      // Skip thin top accent bar re-rendered by master
      if (box.y === 0 && box.height <= 10 && box.width >= dims.width * 0.9) {
        continue;
      }
      const text = extractParagraphText(spXml);
      const spPrBlock = spXml.match(/<p:spPr\b[^>]*>([\s\S]*?)<\/p:spPr>/)?.[1] ?? '';
      const rPrBlock = spXml.match(/<a:rPr\b[^>]*>([\s\S]*?)<\/a:rPr>/)?.[1] ?? '';
      const prstMatch = spPrBlock.match(/<a:prstGeom\b[^>]*prst="([^"]+)"/);
      const prst = prstMatch?.[1] ?? 'rect';
      const hasNoFill = /<a:noFill\b/.test(spPrBlock);
      const fillMatch = hasNoFill
        ? null
        : spPrBlock.match(/<a:solidFill>[\s\S]*?<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
      const textColorMatch = rPrBlock.match(/<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
      const szMatch = spXml.match(/<a:rPr\b[^>]*\bsz="(\d+)"/);
      const boldMatch = spXml.match(/<a:rPr\b[^>]*\bb="1"/);
      const italicMatch = spXml.match(/<a:rPr\b[^>]*\bi="1"/);
      const alignMatch = spXml.match(/<a:pPr\b[^>]*\balgn="(l|ctr|r)"/);
      const latinFontMatch =
        spXml.match(/<a:latin\b[^>]*typeface="([^"]+)"/) ??
        spXml.match(/<a:ea\b[^>]*typeface="([^"]+)"/);
      const normAutofitMatch = spXml.match(/<a:normAutofit\b[^>]*fontScale="(\d+)"/);
      const bodyPrMatch = spXml.match(/<a:bodyPr\b([^>]*)>/);

      if (text.length > 0 && (prst === 'rect' || !fillMatch)) {
        // Scale OOXML point size (1/100 pt) into our 960px canvas coordinates (relative to slide width)
        const widthInches = slideCxEmu / 914400;
        const ptToCanvasPx = dims.width / (widthInches * 72);
        const rawPt = szMatch ? Number(szMatch[1]) / 100 : textIdx === 0 ? 24 : 14;
        const ooxmlScale = normAutofitMatch
          ? Math.max(0.4, Math.min(1, Number(normAutofitMatch[1]) / 100000))
          : 1;
        const requestedFontSize = Math.max(
          10,
          Math.min(64, Math.round(rawPt * ptToCanvasPx * ooxmlScale)),
        );
        const role: TextElement['role'] = requestedFontSize >= 28 ? 'title' : 'body';
        const fontFamily = resolveImportedFontName(
          latinFontMatch?.[1],
          themeFonts,
          role,
          theme,
        );
        const bold = Boolean(boldMatch) || (role === 'title' && !szMatch);
        const italic = Boolean(italicMatch);
        const align: 'left' | 'center' | 'right' =
          alignMatch?.[1] === 'ctr' ? 'center' : alignMatch?.[1] === 'r' ? 'right' : 'left';

        // Parse bodyPr insets if specified
        let padding = 4;
        if (bodyPrMatch) {
          const insMatch = bodyPrMatch[1].match(/\blIns="(\d+)"/);
          if (insMatch) {
            const insPx = Math.round(Number(insMatch[1]) / (slideCxEmu / dims.width));
            padding = Math.max(2, Math.min(14, insPx));
          }
        }

        const baseLineHeight = 1.32;
        const fit = fitTextToSlideBox({
          text,
          width: box.width,
          height: box.height,
          fontSize: requestedFontSize,
          lineHeight: baseLineHeight,
          padding,
          bold,
          italic,
          fontFamily,
          allowAutoShrink: true,
        });

        elements.push({
          id: createSlideUid('txt'),
          type: 'text',
          role,
          x: box.x,
          y: box.y,
          width: box.width,
          height: box.height,
          rotation: 0,
          text,
          fontFamily,
          fontSize: fit.fittedFontSize,
          originalFontSize: fit.wasAutoShrunk ? requestedFontSize : undefined,
          autoShrink: true,
          hasOverflow: fit.stillOverflows || undefined,
          bold,
          italic,
          align,
          color: textColorMatch ? `#${textColorMatch[1]}` : theme.headingColor,
          fill: fillMatch ? `#${fillMatch[1]}` : undefined,
          padding,
          lineHeight: fit.fittedLineHeight,
        });
        textIdx += 1;
      } else if (fillMatch || prst !== 'rect') {
        const shapeKind: ShapeKind =
          prst === 'ellipse'
            ? 'ellipse'
            : prst === 'triangle'
              ? 'triangle'
              : prst === 'rightArrow'
                ? 'arrow-right'
                : prst === 'diamond'
                  ? 'diamond'
                  : prst === 'roundRect'
                    ? 'rounded-rect'
                    : prst === 'line'
                      ? 'line'
                      : 'rect';
        elements.push({
          id: createSlideUid('shp'),
          type: 'shape',
          shapeKind,
          x: box.x,
          y: box.y,
          width: box.width,
          height: box.height,
          rotation: 0,
          fill: fillMatch ? `#${fillMatch[1]}` : theme.primary,
          stroke: 'transparent',
          strokeWidth: 0,
          label: text || undefined,
          labelColor: textColorMatch ? `#${textColorMatch[1]}` : '#FFFFFF',
        });
      }
    }

    // 2. Tables (<p:graphicFrame> containing <a:tbl>)
    const frameRegex = /<p:graphicFrame\b[^>]*>([\s\S]*?)<\/p:graphicFrame>/g;
    let fMatch: RegExpExecArray | null;
    while ((fMatch = frameRegex.exec(xml)) !== null) {
      const fXml = fMatch[1];
      if (!fXml.includes('<a:tbl')) continue;
      const box = parseEmuBox(fXml, slideCxEmu, slideCyEmu, dims.width, dims.height) ?? {
        x: 64,
        y: 120,
        width: 832,
        height: 320,
      };
      const rows: string[][] = [];
      const trRegex = /<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g;
      let trMatch: RegExpExecArray | null;
      while ((trMatch = trRegex.exec(fXml)) !== null) {
        const cells: string[] = [];
        const tcRegex = /<a:tc\b[^>]*>([\s\S]*?)<\/a:tc>/g;
        let tcMatch: RegExpExecArray | null;
        while ((tcMatch = tcRegex.exec(trMatch[1])) !== null) {
          cells.push(extractParagraphText(tcMatch[1]));
        }
        if (cells.length > 0) rows.push(cells);
      }
      if (rows.length > 0) {
        elements.push({
          id: createSlideUid('tbl'),
          type: 'table',
          x: box.x,
          y: box.y,
          width: box.width,
          height: box.height,
          rotation: 0,
          rows,
          headerBg: theme.primary,
          headerColor: '#FFFFFF',
          cellBg: theme.surface,
          cellColor: theme.bodyColor,
          borderColor: '#CBD5E1',
          fontSize: 15,
        });
      }
    }

    // 3. Pictures (<p:pic>)
    const picRegex = /<p:pic\b[^>]*>([\s\S]*?)<\/p:pic>/g;
    let picMatch: RegExpExecArray | null;
    while ((picMatch = picRegex.exec(xml)) !== null) {
      const pXml = picMatch[1];
      const embedMatch = pXml.match(/r:embed="([^"]+)"/);
      if (!embedMatch) continue;
      const target = relMap[embedMatch[1]];
      if (!target) continue;
      const normalized = target.replace(/^\.\.\//, 'ppt/');
      const mediaFile = zip.files[normalized] ?? zip.files[`ppt/slides/${target}`];
      if (!mediaFile) continue;
      const mediaBytes = await mediaFile.async('uint8array');
      const ext = normalized.split('.').pop()?.toLowerCase() ?? 'png';
      const mime =
        ext === 'jpg' || ext === 'jpeg'
          ? 'image/jpeg'
          : ext === 'webp'
            ? 'image/webp'
            : ext === 'gif'
              ? 'image/gif'
              : 'image/png';
      const box = parseEmuBox(pXml, slideCxEmu, slideCyEmu, dims.width, dims.height) ?? {
        x: 120,
        y: 120,
        width: 420,
        height: 280,
      };
      elements.push({
        id: createSlideUid('img'),
        type: 'image',
        x: box.x,
        y: box.y,
        width: box.width,
        height: box.height,
        rotation: 0,
        src: `data:${mime};base64,${uint8ToBase64(mediaBytes)}`,
        alt: `Slide ${slideNum} image`,
      });
    }

    // 4. Speaker Notes
    let notes = '';
    const notesPath = `ppt/notesSlides/notesSlide${slideNum}.xml`;
    if (zip.files[notesPath]) {
      const notesXml = await zip.files[notesPath].async('string');
      const cleaned = notesXml.replace(/<a:fld\b[^>]*type="slidenum"[\s\S]*?<\/a:fld>/g, '');
      notes = extractParagraphText(cleaned);
    }

    parsedSlides.push({
      id: createSlideUid('slide'),
      layout: 'blank',
      background,
      transition: 'fade',
      notes,
      elements,
    });
  }

  return {
    title: cleanTitle,
    themeId: theme.id,
    aspectRatio,
    master: {
      footerText: 'eSZ OFFICE',
      showSlideNumbers: true,
      showAccentBar: true,
    },
    slides: parsedSlides,
    activeSlideId: parsedSlides[0].id,
  };
}

function hexToPdfRgb(hex: string) {
  const clean = stripHexHash(hex, '0F172A');
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  return rgb(r, g, b);
}

function sanitizePdfText(str: string): string {
  return str
    .replace(/•/g, '-')
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/[—–]/g, '-')
    .replace(/[^\x20-\x7E\n]/g, ' ');
}

function wrapTextLines(text: string, maxChars: number): string[] {
  const result: string[] = [];
  for (const para of text.split('\n')) {
    if (!para.trim()) {
      result.push('');
      continue;
    }
    const words = para.split(/\s+/);
    let current = '';
    for (const word of words) {
      const candidate = current ? `${current} ${word}` : word;
      if (candidate.length > maxChars && current) {
        result.push(current);
        current = word;
      } else {
        current = candidate;
      }
    }
    if (current) result.push(current);
  }
  return result;
}

/**
 * Exports a PresentationDeck to a crisp PDF matching its 16:9 or 4:3 aspect ratio using pdf-lib.
 */
export async function deckToPdfBytes(deck: PresentationDeck): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  pdf.setTitle(deck.title || 'Presentation');
  pdf.setProducer('eSZ OFFICE');

  const dims = getSlideDimensions(deck.aspectRatio);
  const slideW = dims.width;
  const slideH = dims.height;

  const fontRegular = await pdf.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const fontItalic = await pdf.embedFont(StandardFonts.HelveticaOblique);
  const theme = getSlideTheme(deck.themeId);

  for (let sIdx = 0; sIdx < deck.slides.length; sIdx += 1) {
    const slide = deck.slides[sIdx];
    const page = pdf.addPage([slideW, slideH]);

    // Background fill
    page.drawRectangle({
      x: 0,
      y: 0,
      width: slideW,
      height: slideH,
      color: hexToPdfRgb(slide.background || '#FFFFFF'),
    });

    if (deck.master.showAccentBar) {
      page.drawRectangle({
        x: 0,
        y: slideH - 6,
        width: slideW,
        height: 6,
        color: hexToPdfRgb(theme.primary),
      });
    }

    for (const el of slide.elements) {
      const pdfY = slideH - el.y - el.height;

      if (el.type === 'shape') {
        if (el.shapeKind === 'ellipse') {
          page.drawEllipse({
            x: el.x + el.width / 2,
            y: pdfY + el.height / 2,
            xScale: el.width / 2,
            yScale: el.height / 2,
            color:
              el.fill && el.fill !== 'transparent' ? hexToPdfRgb(el.fill) : undefined,
            borderColor:
              el.stroke && el.stroke !== 'transparent' ? hexToPdfRgb(el.stroke) : undefined,
            borderWidth: el.strokeWidth || 0,
          });
        } else if (el.shapeKind === 'line') {
          page.drawLine({
            start: {x: el.x, y: pdfY + el.height / 2},
            end: {x: el.x + el.width, y: pdfY + el.height / 2},
            thickness: Math.max(2, el.strokeWidth || 2),
            color: hexToPdfRgb(el.fill || theme.primary),
          });
        } else {
          page.drawRectangle({
            x: el.x,
            y: pdfY,
            width: el.width,
            height: el.height,
            color:
              el.fill && el.fill !== 'transparent' ? hexToPdfRgb(el.fill) : undefined,
            borderColor:
              el.stroke && el.stroke !== 'transparent' ? hexToPdfRgb(el.stroke) : undefined,
            borderWidth: el.strokeWidth || 0,
          });
        }
        if (el.label) {
          const safe = sanitizePdfText(el.label);
          page.drawText(safe, {
            x: el.x + 12,
            y: pdfY + el.height / 2 - 6,
            size: 13,
            font: fontBold,
            color: hexToPdfRgb(el.labelColor || '#FFFFFF'),
          });
        }
      } else if (el.type === 'text') {
        if (el.fill && el.fill !== 'transparent') {
          page.drawRectangle({
            x: el.x,
            y: pdfY,
            width: el.width,
            height: el.height,
            color: hexToPdfRgb(el.fill),
          });
        }
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
        const font = el.bold ? fontBold : el.italic ? fontItalic : fontRegular;
        const size = Math.max(9, fit.fittedFontSize);
        const approxCharW = size * 0.54;
        const maxChars = Math.max(12, Math.floor((el.width - 12) / approxCharW));
        const display = sanitizePdfText(formatListDisplayText(el.text, el.listStyle));
        const lines = wrapTextLines(display, maxChars);
        const lineStep = size * fit.fittedLineHeight;
        let cursorY = slideH - el.y - size - 4;
        for (const line of lines) {
          if (cursorY < pdfY) break;
          if (line.length > 0) {
            page.drawText(line, {
              x: el.x + 4,
              y: cursorY,
              size,
              font,
              color: hexToPdfRgb(el.color || theme.headingColor),
            });
          }
          cursorY -= lineStep;
        }
      } else if (el.type === 'table') {
        const rowCount = Math.max(1, el.rows.length);
        const colCount = Math.max(1, el.rows[0]?.length ?? 1);
        const rowH = el.height / rowCount;
        const colW = el.width / colCount;
        el.rows.forEach((row, rIdx) => {
          row.forEach((cell, cIdx) => {
            const cx = el.x + cIdx * colW;
            const cy = slideH - el.y - (rIdx + 1) * rowH;
            page.drawRectangle({
              x: cx,
              y: cy,
              width: colW,
              height: rowH,
              color: hexToPdfRgb(rIdx === 0 ? el.headerBg : el.cellBg),
              borderColor: hexToPdfRgb(el.borderColor),
              borderWidth: 0.75,
            });
            const safe = sanitizePdfText(cell).slice(0, 32);
            page.drawText(safe, {
              x: cx + 8,
              y: cy + rowH / 2 - 5,
              size: Math.min(13, el.fontSize),
              font: rIdx === 0 ? fontBold : fontRegular,
              color: hexToPdfRgb(rIdx === 0 ? el.headerColor : el.cellColor),
            });
          });
        });
      } else if (el.type === 'chart') {
        page.drawRectangle({
          x: el.x,
          y: pdfY,
          width: el.width,
          height: el.height,
          color: hexToPdfRgb('#FFFFFF'),
          borderColor: hexToPdfRgb('#CBD5E1'),
          borderWidth: 1,
        });
        page.drawText(sanitizePdfText(el.title || 'Chart'), {
          x: el.x + 14,
          y: slideH - el.y - 22,
          size: 13,
          font: fontBold,
          color: hexToPdfRgb('#0F172A'),
        });
        const maxVal = Math.max(1, ...el.values);
        const plotW = el.width - 48;
        const plotH = el.height - 74;
        const barSlot = plotW / Math.max(1, el.values.length);
        el.values.forEach((val, idx) => {
          const bh = Math.max(4, (Math.max(0, val) / maxVal) * plotH);
          const bx = el.x + 24 + idx * barSlot + barSlot * 0.16;
          const bw = Math.max(10, barSlot * 0.68);
          page.drawRectangle({
            x: bx,
            y: pdfY + 30,
            width: bw,
            height: bh,
            color: hexToPdfRgb(el.colors[idx % el.colors.length] || theme.primary),
          });
          page.drawText(sanitizePdfText(el.labels[idx] || `${idx + 1}`).slice(0, 10), {
            x: bx,
            y: pdfY + 12,
            size: 9,
            font: fontRegular,
            color: hexToPdfRgb('#475569'),
          });
        });
      } else if (el.type === 'image' && el.src.startsWith('data:image/')) {
        try {
          const base64 = el.src.split(',')[1] ?? '';
          const raw =
            typeof Buffer !== 'undefined'
              ? Uint8Array.from(Buffer.from(base64, 'base64'))
              : Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
          const embedded = el.src.includes('image/png')
            ? await pdf.embedPng(raw)
            : await pdf.embedJpg(raw);
          page.drawImage(embedded, {
            x: el.x,
            y: pdfY,
            width: el.width,
            height: el.height,
          });
        } catch {
          // Skip unreadable image formats gracefully
        }
      }
    }

    if (deck.master.footerText.trim()) {
      page.drawText(sanitizePdfText(deck.master.footerText), {
        x: 42,
        y: 16,
        size: 10,
        font: fontRegular,
        color: hexToPdfRgb(theme.mutedColor),
      });
    }
    if (deck.master.showSlideNumbers) {
      page.drawText(`${sIdx + 1} / ${deck.slides.length}`, {
        x: slideW - 76,
        y: 16,
        size: 10,
        font: fontRegular,
        color: hexToPdfRgb(theme.mutedColor),
      });
    }
  }

  return pdf.save();
}

export function buildPersistedSlidesMeta(deck: PresentationDeck): PersistedSlidesMeta {
  return {
    version: 1,
    deck,
  };
}

export function buildDeckPreviewText(deck: PresentationDeck): string {
  return deck.slides
    .map((s, idx) => {
      const heading = s.elements.find((e): e is TextElement => e.type === 'text');
      return `Slide ${idx + 1}: ${heading?.text.split('\n')[0] || 'Untitled'}`;
    })
    .join(' | ')
    .slice(0, 240);
}

export function createNewElement(
  kind: 'text' | 'shape' | 'table' | 'chart' | 'image',
  theme: SlideTheme,
  opts?: {
    shapeKind?: ShapeKind;
    chartKind?: SlideChartKind;
    chartTitle?: string;
    chartLabels?: string[];
    chartValues?: number[];
    tableRows?: number;
    tableCols?: number;
    imageSrc?: string;
    imageAlt?: string;
  },
): SlideElement {
  if (kind === 'text') {
    const el: TextElement = {
      id: createSlideUid('txt'),
      type: 'text',
      role: 'body',
      x: 140,
      y: 170,
      width: 440,
      height: 110,
      rotation: 0,
      text: 'Double-click to edit text',
      fontFamily: theme.bodyFont,
      fontSize: 22,
      align: 'left',
      color: theme.headingColor,
      lineHeight: 1.4,
    };
    return el;
  }

  if (kind === 'shape') {
    const sk = opts?.shapeKind ?? 'rounded-rect';
    const el: ShapeElement = {
      id: createSlideUid('shp'),
      type: 'shape',
      shapeKind: sk,
      x: 260,
      y: 160,
      width: sk === 'line' ? 280 : 220,
      height: sk === 'line' ? 16 : 130,
      rotation: 0,
      fill: theme.primary,
      stroke: 'transparent',
      strokeWidth: sk === 'line' ? 3 : 0,
      cornerRadius: sk === 'rounded-rect' ? 14 : 0,
      label: sk === 'line' ? undefined : 'Callout',
      labelColor: '#FFFFFF',
    };
    return el;
  }

  if (kind === 'table') {
    const rCount = Math.max(2, Math.min(8, opts?.tableRows ?? 4));
    const cCount = Math.max(2, Math.min(6, opts?.tableCols ?? 3));
    const rows: string[][] = Array.from({length: rCount}, (_, r) =>
      Array.from({length: cCount}, (_, c) =>
        r === 0 ? `Column ${c + 1}` : `Row ${r}, Col ${c + 1}`,
      ),
    );
    const el: TableElement = {
      id: createSlideUid('tbl'),
      type: 'table',
      x: 110,
      y: 130,
      width: 720,
      height: Math.min(340, rCount * 58),
      rotation: 0,
      rows,
      headerBg: theme.primary,
      headerColor: '#FFFFFF',
      cellBg: theme.surface,
      cellColor: theme.bodyColor,
      borderColor: '#CBD5E1',
      fontSize: 15,
    };
    return el;
  }

  if (kind === 'chart') {
    const el: ChartElement = {
      id: createSlideUid('cht'),
      type: 'chart',
      chartKind: opts?.chartKind ?? 'bar',
      x: 160,
      y: 120,
      width: 520,
      height: 320,
      rotation: 0,
      title: opts?.chartTitle || 'Quarterly Results',
      labels: opts?.chartLabels ?? ['Q1', 'Q2', 'Q3', 'Q4'],
      values: opts?.chartValues ?? [35, 58, 74, 96],
      colors: theme.chartPalette,
    };
    return el;
  }

  const img: ImageElement = {
    id: createSlideUid('img'),
    type: 'image',
    x: 200,
    y: 110,
    width: 440,
    height: 290,
    rotation: 0,
    src: opts?.imageSrc || '',
    alt: opts?.imageAlt || 'Slide image',
    cornerRadius: 8,
  };
  return img;
}
