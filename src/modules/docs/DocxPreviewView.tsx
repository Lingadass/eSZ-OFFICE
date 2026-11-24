import React, {useEffect, useRef, useState} from 'react';
import {resolveRenderFontStack} from './docFonts';
import {
  hydrateDocxPreviewDom,
  preprocessDocxForPreview,
  type DocxDocDefaults,
} from './docxPreviewUtils';

export interface DocxPreviewMetrics {
  pageCount: number;
  /** Unscaled offsetTop of each rendered section.esz-docx relative to the preview wrapper */
  pageTops: number[];
  /** Unscaled offsetHeight of each rendered section.esz-docx */
  pageHeights: number[];
  wrapperWidth: number;
  wrapperHeight: number;
}

interface DocxPreviewViewProps {
  /** Raw unmodified .docx bytes from device storage (never mutated) */
  buffer: ArrayBuffer;
  zoom: number;
  onMetricsChange: (metrics: DocxPreviewMetrics) => void;
  onError?: (error: Error) => void;
  previewRootRef?: React.RefObject<HTMLDivElement | null>;
}

/**
 * Faithful Microsoft Word View Mode canvas powered by docx-preview (Apache-2.0).
 * - Preserves original stored .docx bytes untouched
 * - Pre-resolves mc:AlternateContent using mc:Choice (dropping mc:Fallback)
 * - Renders headers, footers, footnotes, endnotes, tables, page borders, and images
 * - Zooms purely via CSS transform without re-layout
 */
export const DocxPreviewView: React.FC<DocxPreviewViewProps> = ({
  buffer,
  zoom,
  onMetricsChange,
  onError,
  previewRootRef,
}) => {
  const styleHostRef = useRef<HTMLDivElement>(null);
  const bodyHostRef = useRef<HTMLDivElement>(null);
  const onMetricsChangeRef = useRef(onMetricsChange);
  const onErrorRef = useRef(onError);
  onMetricsChangeRef.current = onMetricsChange;
  onErrorRef.current = onError;

  const [rendering, setRendering] = useState(true);
  const [naturalSize, setNaturalSize] = useState<{width: number; height: number}>({
    width: 816,
    height: 1056,
  });
  const [docDefaults, setDocDefaults] = useState<DocxDocDefaults>({
    fontFamily: null,
    fontSizePt: null,
    color: null,
  });

  useEffect(() => {
    let cancelled = false;
    let resizeObserver: ResizeObserver | null = null;
    const imgCleanupFns: Array<() => void> = [];

    const measureUnscaled = () => {
      const bodyHost = bodyHostRef.current;
      if (!bodyHost || cancelled) return;

      const wrapper =
        bodyHost.querySelector<HTMLElement>('.esz-docx-wrapper') ?? bodyHost;
      const sections = Array.from(bodyHost.querySelectorAll<HTMLElement>('section.esz-docx'));

      // offsetWidth / offsetHeight / offsetTop are strictly unscaled CSS layout pixels
      // even when an ancestor carries `transform: scale(zoom)`.
      let maxSectionWidth = 0;
      const pageTops: number[] = [];
      const pageHeights: number[] = [];

      for (const sec of sections) {
        maxSectionWidth = Math.max(maxSectionWidth, sec.offsetWidth);
        pageTops.push(sec.offsetTop);
        pageHeights.push(sec.offsetHeight);
      }

      const wrapperWidth = Math.max(
        320,
        Math.ceil(Math.max(wrapper.scrollWidth, wrapper.offsetWidth, maxSectionWidth)),
      );
      const wrapperHeight = Math.max(
        200,
        Math.ceil(Math.max(wrapper.scrollHeight, wrapper.offsetHeight)),
      );

      setNaturalSize((prev) =>
        prev.width === wrapperWidth && prev.height === wrapperHeight
          ? prev
          : {width: wrapperWidth, height: wrapperHeight},
      );

      onMetricsChangeRef.current({
        pageCount: Math.max(1, sections.length),
        pageTops,
        pageHeights,
        wrapperWidth,
        wrapperHeight,
      });
    };

    (async () => {
      const bodyHost = bodyHostRef.current;
      const styleHost = styleHostRef.current;
      if (!bodyHost || !styleHost) return;

      setRendering(true);
      bodyHost.innerHTML = '';
      styleHost.innerHTML = '';

      try {
        const [{renderAsync}, preprocessed] = await Promise.all([
          import('docx-preview'),
          preprocessDocxForPreview(buffer),
        ]);
        if (cancelled) return;

        setDocDefaults(preprocessed.docDefaults);

        await renderAsync(preprocessed.buffer, bodyHost, styleHost, {
          className: 'esz-docx',
          inWrapper: true,
          hideWrapperOnPrint: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          ignoreLastRenderedPageBreak: true,
          experimental: false,
          trimXmlDeclaration: true,
          useBase64URL: true,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
          renderChanges: false,
          renderComments: false,
        });
        if (cancelled) return;

        hydrateDocxPreviewDom(bodyHost, preprocessed, styleHost);
        setRendering(false);

        // Measure immediately after paint
        window.requestAnimationFrame(() => {
          measureUnscaled();
          window.requestAnimationFrame(measureUnscaled);
        });

        // Re-measure after webfonts finish loading
        if (typeof document !== 'undefined' && document.fonts?.ready) {
          void document.fonts.ready.then(() => {
            if (!cancelled) window.requestAnimationFrame(measureUnscaled);
          });
        }

        // Re-measure when embedded images finish decoding
        const imgs = Array.from(bodyHost.querySelectorAll('img'));
        for (const img of imgs) {
          if (!img.complete) {
            const onLoadOrErr = () => {
              if (!cancelled) window.requestAnimationFrame(measureUnscaled);
            };
            img.addEventListener('load', onLoadOrErr, {once: true});
            img.addEventListener('error', onLoadOrErr, {once: true});
            imgCleanupFns.push(() => {
              img.removeEventListener('load', onLoadOrErr);
              img.removeEventListener('error', onLoadOrErr);
            });
          }
        }

        // Keep outer zoom wrapper synced with any async layout updates
        const wrapper = bodyHost.querySelector<HTMLElement>('.esz-docx-wrapper') ?? bodyHost;
        if (typeof ResizeObserver !== 'undefined') {
          resizeObserver = new ResizeObserver(() => {
            if (!cancelled) measureUnscaled();
          });
          resizeObserver.observe(wrapper);
        }
      } catch (err) {
        if (!cancelled) {
          setRendering(false);
          onErrorRef.current?.(
            err instanceof Error ? err : new Error('Failed to render Word document preview.'),
          );
        }
      }
    })();

    return () => {
      cancelled = true;
      resizeObserver?.disconnect();
      for (const cleanup of imgCleanupFns) cleanup();
    };
  }, [buffer]);

  const scaledWidth = Math.ceil(naturalSize.width * zoom);
  const scaledHeight = Math.ceil(naturalSize.height * zoom);
  const defaultDocFont = resolveRenderFontStack(docDefaults.fontFamily, 'Times New Roman');

  return (
    <div
      data-testid="docx-preview-root"
      className="flex min-w-full w-max flex-col items-center pb-8"
    >
      {rendering && (
        <div
          data-testid="docx-preview-loading"
          className="mx-auto my-6 w-full max-w-3xl space-y-4 rounded-xl bg-white p-10 shadow-sm dark:bg-slate-900"
        >
          <div className="h-7 w-1/2 animate-pulse rounded bg-slate-200 dark:bg-slate-800" />
          <div className="h-4 w-full animate-pulse rounded bg-slate-100 dark:bg-slate-800" />
          <div className="h-4 w-5/6 animate-pulse rounded bg-slate-100 dark:bg-slate-800" />
          <div className="h-4 w-4/6 animate-pulse rounded bg-slate-100 dark:bg-slate-800" />
        </div>
      )}

      {/* Scoped style container populated by docx-preview */}
      <div ref={styleHostRef} className="hidden" aria-hidden="true" />

      {/* Outer zoom-sized container so scrollbars match scaled dimensions without re-layout */}
      <div
        data-testid="docx-preview-zoom-shell"
        style={{
          width: `${scaledWidth}px`,
          height: rendering ? '0px' : `${scaledHeight}px`,
          overflow: rendering ? 'hidden' : 'visible',
        }}
        className="relative shrink-0"
      >
        {/* Unscaled 1:1 docx-preview stage transformed purely by zoom */}
        <div
          ref={(el) => {
            (bodyHostRef as React.MutableRefObject<HTMLDivElement | null>).current = el;
            if (previewRootRef) {
              (previewRootRef as React.MutableRefObject<HTMLDivElement | null>).current = el;
            }
          }}
          data-testid="docx-preview-stage"
          className="esz-docx-preview-host origin-top-left"
          style={
            {
              width: `${naturalSize.width}px`,
              transform: `scale(${zoom})`,
              '--esz-doc-default-font': defaultDocFont,
              '--esz-doc-default-size': docDefaults.fontSizePt
                ? `${docDefaults.fontSizePt}pt`
                : '11pt',
              '--esz-doc-default-color': docDefaults.color ?? '#000000',
            } as React.CSSProperties
          }
        />
      </div>
    </div>
  );
};
