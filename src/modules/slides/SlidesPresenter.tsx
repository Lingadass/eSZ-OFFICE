import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import {
  ChevronLeft,
  ChevronRight,
  Clock,
  Crosshair,
  Maximize2,
  MessageSquareText,
  MonitorOff,
  Pause,
  Play,
  RotateCcw,
  X,
} from 'lucide-react';
import {SlideSvgRenderer} from './SlidesCanvas';
import {getSlideDimensions} from './slidesTextFit';
import type {PresentationDeck} from './slidesTypes';

interface SlidesPresenterProps {
  deck: PresentationDeck;
  initialIndex: number;
  onExit: () => void;
}

function formatTimer(seconds: number): string {
  const m = Math.floor(seconds / 60);
  const s = seconds % 60;
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}

export const SlidesPresenter: React.FC<SlidesPresenterProps> = ({
  deck,
  initialIndex,
  onExit,
}) => {
  const [index, setIndex] = useState(
    Math.max(0, Math.min(deck.slides.length - 1, initialIndex)),
  );
  const [showPresenterPane, setShowPresenterPane] = useState(true);
  const [blackScreen, setBlackScreen] = useState(false);
  const [laserEnabled, setLaserEnabled] = useState(false);
  const [laserPos, setLaserPos] = useState<{x: number; y: number} | null>(null);
  const [elapsedSec, setElapsedSec] = useState(0);
  const [timerRunning, setTimerRunning] = useState(true);
  const [animKey, setAnimKey] = useState(0);

  const stageHostRef = useRef<HTMLDivElement | null>(null);
  const [stageBox, setStageBox] = useState<{width: number; height: number} | null>(null);

  // Touch swipe tracking (left = next slide, right = previous slide)
  const touchStartRef = useRef<{x: number; y: number} | null>(null);
  const swipedRecentlyRef = useRef(false);

  const dims = useMemo(() => getSlideDimensions(deck.aspectRatio), [deck.aspectRatio]);
  const currentSlide = deck.slides[index] ?? deck.slides[0];
  const nextSlide = deck.slides[index + 1] ?? null;

  const goNext = useCallback(() => {
    setBlackScreen(false);
    setIndex((i) => {
      const next = Math.min(deck.slides.length - 1, i + 1);
      if (next !== i) setAnimKey((k) => k + 1);
      return next;
    });
  }, [deck.slides.length]);

  const goPrev = useCallback(() => {
    setBlackScreen(false);
    setIndex((i) => {
      const prev = Math.max(0, i - 1);
      if (prev !== i) setAnimKey((k) => k + 1);
      return prev;
    });
  }, []);

  // Maximize slide inside presenter viewport while honoring exact 16:9 or 4:3 aspect ratio
  useEffect(() => {
    const host = stageHostRef.current;
    if (!host) return;
    const measure = () => {
      const rect = host.getBoundingClientRect();
      const pad = rect.width < 640 ? 24 : 48;
      const availW = Math.max(200, rect.width - pad);
      const availH = Math.max(150, rect.height - pad);
      const scale = Math.min(availW / dims.width, availH / dims.height);
      setStageBox({
        width: Math.max(160, Math.round(dims.width * scale)),
        height: Math.max(120, Math.round(dims.height * scale)),
      });
    };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(host);
    return () => ro.disconnect();
  }, [dims.width, dims.height, showPresenterPane]);

  useEffect(() => {
    if (!timerRunning) return;
    const id = window.setInterval(() => setElapsedSec((s) => s + 1), 1000);
    return () => window.clearInterval(id);
  }, [timerRunning]);

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        if (document.fullscreenElement) {
          void document.exitFullscreen().catch(() => {});
        }
        onExit();
      } else if (
        e.key === 'ArrowRight' ||
        e.key === 'ArrowDown' ||
        e.key === ' ' ||
        e.key === 'PageDown'
      ) {
        e.preventDefault();
        e.stopPropagation();
        goNext();
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') {
        e.preventDefault();
        e.stopPropagation();
        goPrev();
      } else if (e.key === 'Home') {
        e.preventDefault();
        setIndex(0);
        setAnimKey((k) => k + 1);
      } else if (e.key === 'End') {
        e.preventDefault();
        setIndex(deck.slides.length - 1);
        setAnimKey((k) => k + 1);
      } else if (e.key.toLowerCase() === 'b') {
        e.preventDefault();
        setBlackScreen((b) => !b);
      } else if (e.key.toLowerCase() === 'n') {
        e.preventDefault();
        setShowPresenterPane((p) => !p);
      }
    };

    // Capture phase ensures Esc closes Presenter cleanly without triggering outer workspace shortcuts
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, [deck.slides.length, goNext, goPrev, onExit]);

  const requestFullscreen = () => {
    if (document.fullscreenElement) {
      void document.exitFullscreen().catch(() => {});
    } else {
      void document.documentElement.requestFullscreen?.().catch(() => {});
    }
  };

  const handleTouchStart = (e: React.TouchEvent<HTMLDivElement>) => {
    if (e.touches.length !== 1) return;
    const t = e.touches[0];
    touchStartRef.current = {x: t.clientX, y: t.clientY};
  };

  const handleTouchEnd = (e: React.TouchEvent<HTMLDivElement>) => {
    const start = touchStartRef.current;
    touchStartRef.current = null;
    if (!start || e.changedTouches.length === 0) return;
    const t = e.changedTouches[0];
    const dx = t.clientX - start.x;
    const dy = t.clientY - start.y;
    if (Math.abs(dx) >= 45 && Math.abs(dx) > Math.abs(dy) * 1.2) {
      swipedRecentlyRef.current = true;
      window.setTimeout(() => {
        swipedRecentlyRef.current = false;
      }, 300);
      if (dx < 0) {
        goNext();
      } else {
        goPrev();
      }
    }
  };

  const transitionAnimClass =
    currentSlide.transition === 'fade'
      ? 'animate-[fadeIn_260ms_ease-out]'
      : currentSlide.transition === 'zoom'
        ? 'animate-[zoomIn_260ms_ease-out]'
        : currentSlide.transition === 'slide'
          ? 'animate-[slideInRight_260ms_ease-out]'
          : '';

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Presentation slide show"
      data-testid="slides-presenter"
      data-aspect-ratio={dims.aspectRatio}
      className="fixed inset-0 z-50 flex flex-col bg-slate-950 text-slate-100"
    >
      {/* Top Presenter Controls */}
      <header className="flex shrink-0 items-center justify-between gap-3 border-b border-slate-800 bg-slate-900/95 px-4 py-2">
        <div className="flex min-w-0 items-center gap-3">
          <span className="truncate text-sm font-semibold text-white">{deck.title}</span>
          <span
            data-testid="slides-presenter-counter"
            className="font-mono text-xs tabular-nums text-slate-400"
          >
            Slide {index + 1} / {deck.slides.length}
          </span>
          <span className="hidden rounded border border-slate-700 px-1.5 py-0.5 font-mono text-[10px] text-slate-400 sm:inline">
            {dims.aspectRatio}
          </span>
        </div>

        <div className="flex shrink-0 items-center gap-1.5">
          <div className="hidden items-center gap-1.5 rounded-md border border-slate-700 bg-slate-950 px-2.5 py-1 text-xs font-mono tabular-nums text-slate-200 sm:flex">
            <Clock className="h-3.5 w-3.5 text-orange-400" />
            <span>{formatTimer(elapsedSec)}</span>
            <button
              type="button"
              onClick={() => setTimerRunning((r) => !r)}
              aria-label={timerRunning ? 'Pause presenter timer' : 'Resume presenter timer'}
              className="ml-1 rounded p-0.5 text-slate-400 hover:text-white"
            >
              {timerRunning ? <Pause className="h-3 w-3" /> : <Play className="h-3 w-3" />}
            </button>
            <button
              type="button"
              onClick={() => setElapsedSec(0)}
              aria-label="Reset presenter timer"
              className="rounded p-0.5 text-slate-400 hover:text-white"
            >
              <RotateCcw className="h-3 w-3" />
            </button>
          </div>

          <button
            type="button"
            onClick={() => setLaserEnabled((l) => !l)}
            aria-pressed={laserEnabled}
            aria-label="Toggle laser pointer"
            className={`inline-flex h-8 items-center gap-1.5 whitespace-nowrap rounded-md border px-2.5 text-xs font-medium ${
              laserEnabled
                ? 'border-rose-500 bg-rose-950/60 text-rose-200'
                : 'border-slate-700 bg-slate-800 text-slate-200 hover:bg-slate-700'
            }`}
          >
            <Crosshair className="h-3.5 w-3.5" />
            <span className="hidden md:inline">Laser</span>
          </button>

          <button
            type="button"
            onClick={() => setBlackScreen((b) => !b)}
            aria-pressed={blackScreen}
            aria-label="Toggle black screen (B)"
            className={`inline-flex h-8 items-center gap-1.5 whitespace-nowrap rounded-md border px-2.5 text-xs font-medium ${
              blackScreen
                ? 'border-amber-500 bg-amber-950/60 text-amber-200'
                : 'border-slate-700 bg-slate-800 text-slate-200 hover:bg-slate-700'
            }`}
          >
            <MonitorOff className="h-3.5 w-3.5" />
            <span className="hidden md:inline">Blackout</span>
          </button>

          <button
            type="button"
            onClick={() => setShowPresenterPane((p) => !p)}
            aria-pressed={showPresenterPane}
            aria-label="Toggle presenter notes sidebar (N)"
            className={`inline-flex h-8 items-center gap-1.5 whitespace-nowrap rounded-md border px-2.5 text-xs font-medium ${
              showPresenterPane
                ? 'border-orange-500 bg-orange-950/50 text-orange-200'
                : 'border-slate-700 bg-slate-800 text-slate-200 hover:bg-slate-700'
            }`}
          >
            <MessageSquareText className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Notes</span>
          </button>

          <button
            type="button"
            onClick={requestFullscreen}
            aria-label="Toggle browser fullscreen"
            className="inline-flex h-8 w-8 items-center justify-center rounded-md border border-slate-700 bg-slate-800 text-slate-200 hover:bg-slate-700"
          >
            <Maximize2 className="h-3.5 w-3.5" />
          </button>

          <button
            type="button"
            onClick={onExit}
            aria-label="Exit slide show (Esc)"
            className="inline-flex h-8 items-center gap-1 whitespace-nowrap rounded-md bg-rose-600 px-3 text-xs font-semibold text-white hover:bg-rose-500"
          >
            <X className="h-3.5 w-3.5" />
            Exit
          </button>
        </div>
      </header>

      {/* Main Stage + Optional Presenter Sidebar */}
      <div className="flex min-h-0 flex-1 overflow-hidden">
        <div
          ref={stageHostRef}
          data-testid="slides-presenter-stage"
          className="relative flex min-w-0 flex-1 touch-pan-y items-center justify-center p-3 sm:p-6"
          onTouchStart={handleTouchStart}
          onTouchEnd={handleTouchEnd}
          onMouseMove={(e) => {
            if (!laserEnabled) return;
            const rect = e.currentTarget.getBoundingClientRect();
            setLaserPos({x: e.clientX - rect.left, y: e.clientY - rect.top});
          }}
          onMouseLeave={() => setLaserPos(null)}
          onClick={() => {
            if (swipedRecentlyRef.current) return;
            if (blackScreen) {
              setBlackScreen(false);
              return;
            }
            goNext();
          }}
        >
          {blackScreen ? (
            <div className="flex flex-col items-center gap-2 text-center text-slate-500">
              <p className="text-sm font-medium">Screen blanked</p>
              <p className="text-xs">Press B or click anywhere to resume presentation</p>
            </div>
          ) : (
            <div
              key={`${currentSlide.id}-${animKey}`}
              data-testid="slides-presenter-frame"
              data-aspect-ratio={dims.aspectRatio}
              style={
                stageBox
                  ? {width: stageBox.width, height: stageBox.height}
                  : undefined
              }
              className={`${stageBox ? '' : `${dims.cssAspect} w-full max-w-6xl`} overflow-hidden rounded-lg bg-white shadow-2xl ring-1 ring-white/10 transition-all duration-200 ${transitionAnimClass}`}
            >
              <SlideSvgRenderer
                slide={currentSlide}
                slideIndex={index}
                slideCount={deck.slides.length}
                themeId={deck.themeId}
                aspectRatio={deck.aspectRatio}
                master={deck.master}
              />
            </div>
          )}

          {laserEnabled && laserPos && (
            <div
              aria-hidden="true"
              className="pointer-events-none absolute h-3.5 w-3.5 -translate-x-1/2 -translate-y-1/2 rounded-full bg-rose-500 shadow-[0_0_12px_4px_rgba(244,63,94,0.85)]"
              style={{left: laserPos.x, top: laserPos.y}}
            />
          )}
        </div>

        {showPresenterPane && (
          <aside
            aria-label="Presenter view notes and next slide"
            className="hidden w-80 shrink-0 flex-col border-l border-slate-800 bg-slate-900 p-4 lg:flex"
          >
            <div className="mb-3">
              <div className="mb-1.5 text-xs font-semibold text-slate-400">Next Slide</div>
              {nextSlide ? (
                <div className={`${dims.cssAspect} overflow-hidden rounded-md border border-slate-700 bg-white`}>
                  <SlideSvgRenderer
                    slide={nextSlide}
                    slideIndex={index + 1}
                    slideCount={deck.slides.length}
                    themeId={deck.themeId}
                    aspectRatio={deck.aspectRatio}
                    master={deck.master}
                  />
                </div>
              ) : (
                <div className={`flex ${dims.cssAspect} items-center justify-center rounded-md border border-dashed border-slate-700 text-xs text-slate-500`}>
                  End of presentation
                </div>
              )}
            </div>

            <div className="flex min-h-0 flex-1 flex-col border-t border-slate-800 pt-3">
              <div className="mb-1.5 text-xs font-semibold text-slate-400">Speaker Notes</div>
              <div className="flex-1 overflow-y-auto rounded-md border border-slate-800 bg-slate-950 p-3 text-sm leading-relaxed text-slate-200 whitespace-pre-wrap">
                {currentSlide.notes.trim() || 'No speaker notes for this slide.'}
              </div>
            </div>
          </aside>
        )}
      </div>

      {/* Bottom Navigation Bar */}
      <footer className="flex shrink-0 items-center justify-between border-t border-slate-800 bg-slate-900 px-4 py-2 text-xs">
        <span className="text-slate-400">
          Arrow keys / Space / Swipe to navigate · B to blackout · Esc to exit
        </span>
        <div className="flex items-center gap-2">
          <button
            type="button"
            disabled={index === 0}
            onClick={goPrev}
            aria-label="Previous slide"
            className="inline-flex h-8 items-center gap-1 rounded-md border border-slate-700 bg-slate-800 px-3 font-medium text-slate-100 hover:bg-slate-700 disabled:opacity-40"
          >
            <ChevronLeft className="h-4 w-4" />
            Prev
          </button>
          <button
            type="button"
            disabled={index >= deck.slides.length - 1}
            onClick={goNext}
            aria-label="Next slide"
            className="inline-flex h-8 items-center gap-1 rounded-md bg-orange-600 px-3 font-semibold text-white hover:bg-orange-500 disabled:opacity-40"
          >
            Next
            <ChevronRight className="h-4 w-4" />
          </button>
        </div>
      </footer>
    </div>
  );
};
