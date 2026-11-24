import React, {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
} from 'react';
import Konva from 'konva';
import {
  Arrow,
  Circle,
  Ellipse,
  Group,
  Image as KonvaImage,
  Layer,
  Line,
  Rect,
  RegularPolygon,
  Stage,
  Text,
  Transformer,
  Wedge,
} from 'react-konva';
import {formatListDisplayText, getSlideTheme} from './slidesPure';
import {
  fitTextToSlideBox,
  getSlideDimensions,
  getSlideFontCssStack,
} from './slidesTextFit';
import type {
  ChartElement,
  ImageElement,
  ShapeElement,
  SlideAspectRatio,
  SlideElement,
  SlideItem,
  SlideMasterSettings,
  TableElement,
  TextElement,
} from './slidesTypes';

export interface SlidesCanvasHandle {
  exportPngDataUrl: () => string | null;
}

interface SlidesCanvasProps {
  slide: SlideItem;
  slideIndex: number;
  slideCount: number;
  themeId: string;
  aspectRatio?: SlideAspectRatio;
  master: SlideMasterSettings;
  selectedElementId: string | null;
  zoom: number;
  readOnly?: boolean;
  onSelectElement: (id: string | null) => void;
  onUpdateElement: (id: string, patch: Partial<SlideElement>) => void;
}

const LoadedKonvaImage: React.FC<{
  element: ImageElement;
  readOnly?: boolean;
  onSelect: () => void;
  onDragEnd: (x: number, y: number) => void;
  onTransformEnd: (node: Konva.Node) => void;
  setRef: (node: Konva.Node | null) => void;
}> = ({element, readOnly = false, onSelect, onDragEnd, onTransformEnd, setRef}) => {
  const [img, setImg] = useState<HTMLImageElement | null>(null);

  useEffect(() => {
    if (!element.src) {
      setImg(null);
      return;
    }
    const imageObj = new window.Image();
    imageObj.crossOrigin = 'anonymous';
    imageObj.onload = () => setImg(imageObj);
    imageObj.src = element.src;
  }, [element.src]);

  if (!img) {
    return (
      <Rect
        ref={(n) => setRef(n)}
        id={element.id}
        x={element.x}
        y={element.y}
        width={element.width}
        height={element.height}
        rotation={element.rotation}
        fill="#E2E8F0"
        stroke="#94A3B8"
        strokeWidth={1}
        dash={[6, 4]}
        cornerRadius={element.cornerRadius ?? 8}
        draggable={!readOnly}
        onClick={readOnly ? undefined : onSelect}
        onTap={readOnly ? undefined : onSelect}
        onDragEnd={(e) => {
          if (!readOnly) onDragEnd(e.target.x(), e.target.y());
        }}
        onTransformEnd={(e) => {
          if (!readOnly) onTransformEnd(e.target);
        }}
      />
    );
  }

  return (
    <KonvaImage
      ref={(n) => setRef(n)}
      id={element.id}
      image={img}
      x={element.x}
      y={element.y}
      width={element.width}
      height={element.height}
      rotation={element.rotation}
      cornerRadius={element.cornerRadius ?? 0}
      opacity={element.opacity ?? 1}
      draggable={!readOnly}
      onClick={readOnly ? undefined : onSelect}
      onTap={readOnly ? undefined : onSelect}
      onDragEnd={(e) => {
        if (!readOnly) onDragEnd(e.target.x(), e.target.y());
      }}
      onTransformEnd={(e) => {
        if (!readOnly) onTransformEnd(e.target);
      }}
    />
  );
};

const KonvaTableNode: React.FC<{
  element: TableElement;
  readOnly?: boolean;
  onSelect: () => void;
  onDragEnd: (x: number, y: number) => void;
  onTransformEnd: (node: Konva.Node) => void;
  setRef: (node: Konva.Node | null) => void;
}> = ({element, readOnly = false, onSelect, onDragEnd, onTransformEnd, setRef}) => {
  const rowCount = Math.max(1, element.rows.length);
  const colCount = Math.max(1, element.rows[0]?.length ?? 1);
  const rowH = element.height / rowCount;
  const colW = element.width / colCount;

  return (
    <Group
      ref={(n) => setRef(n)}
      id={element.id}
      x={element.x}
      y={element.y}
      width={element.width}
      height={element.height}
      rotation={element.rotation}
      draggable={!readOnly}
      onClick={readOnly ? undefined : onSelect}
      onTap={readOnly ? undefined : onSelect}
      onDragEnd={(e) => {
        if (!readOnly) onDragEnd(e.target.x(), e.target.y());
      }}
      onTransformEnd={(e) => {
        if (!readOnly) onTransformEnd(e.target);
      }}
    >
      {element.rows.map((row, rIdx) =>
        row.map((cell, cIdx) => (
          <Group key={`${rIdx}-${cIdx}`} x={cIdx * colW} y={rIdx * rowH}>
            <Rect
              width={colW}
              height={rowH}
              fill={rIdx === 0 ? element.headerBg : element.cellBg}
              stroke={element.borderColor}
              strokeWidth={1}
            />
            <Text
              x={8}
              y={6}
              width={Math.max(12, colW - 16)}
              height={Math.max(12, rowH - 12)}
              text={cell}
              fontSize={element.fontSize}
              fontFamily="DM Sans, sans-serif"
              fontStyle={rIdx === 0 ? 'bold' : 'normal'}
              fill={rIdx === 0 ? element.headerColor : element.cellColor}
              verticalAlign="middle"
              ellipsis
            />
          </Group>
        )),
      )}
    </Group>
  );
};

const KonvaChartNode: React.FC<{
  element: ChartElement;
  readOnly?: boolean;
  onSelect: () => void;
  onDragEnd: (x: number, y: number) => void;
  onTransformEnd: (node: Konva.Node) => void;
  setRef: (node: Konva.Node | null) => void;
}> = ({element, readOnly = false, onSelect, onDragEnd, onTransformEnd, setRef}) => {
  const w = element.width;
  const h = element.height;
  const padL = 42;
  const padR = 20;
  const padT = 44;
  const padB = 34;
  const plotW = Math.max(40, w - padL - padR);
  const plotH = Math.max(40, h - padT - padB);
  const vals = element.values.length > 0 ? element.values : [10, 20, 30];
  const labels = element.labels.length > 0 ? element.labels : vals.map((_, i) => `Item ${i + 1}`);
  const maxVal = Math.max(1, ...vals);
  const totalVal = Math.max(1, vals.reduce((a, b) => a + Math.max(0, b), 0));

  const linePoints: number[] = [];
  if (element.chartKind === 'line') {
    vals.forEach((v, i) => {
      const px = padL + ((i + 0.5) / vals.length) * plotW;
      const py = padT + plotH - (Math.max(0, v) / maxVal) * plotH;
      linePoints.push(px, py);
    });
  }

  let pieStart = -90;

  return (
    <Group
      ref={(n) => setRef(n)}
      id={element.id}
      x={element.x}
      y={element.y}
      width={w}
      height={h}
      rotation={element.rotation}
      draggable={!readOnly}
      onClick={readOnly ? undefined : onSelect}
      onTap={readOnly ? undefined : onSelect}
      onDragEnd={(e) => {
        if (!readOnly) onDragEnd(e.target.x(), e.target.y());
      }}
      onTransformEnd={(e) => {
        if (!readOnly) onTransformEnd(e.target);
      }}
    >
      <Rect
        width={w}
        height={h}
        fill="#FFFFFF"
        stroke="#E2E8F0"
        strokeWidth={1.2}
        cornerRadius={10}
      />
      <Text
        x={16}
        y={12}
        width={w - 32}
        text={element.title || 'Chart'}
        fontSize={15}
        fontStyle="bold"
        fontFamily="Plus Jakarta Sans, sans-serif"
        fill="#0F172A"
      />

      {element.chartKind !== 'pie' && (
        <>
          {[0, 0.25, 0.5, 0.75, 1].map((t) => {
            const gy = padT + plotH * (1 - t);
            return (
              <Group key={t}>
                <Line
                  points={[padL, gy, padL + plotW, gy]}
                  stroke="#E2E8F0"
                  strokeWidth={1}
                />
                <Text
                  x={4}
                  y={gy - 6}
                  width={padL - 10}
                  align="right"
                  text={String(Math.round(maxVal * t))}
                  fontSize={10}
                  fill="#64748B"
                />
              </Group>
            );
          })}
        </>
      )}

      {element.chartKind === 'bar' &&
        vals.map((v, i) => {
          const slot = plotW / vals.length;
          const bw = Math.max(12, slot * 0.56);
          const bx = padL + i * slot + (slot - bw) / 2;
          const bh = Math.max(2, (Math.max(0, v) / maxVal) * plotH);
          const color = element.colors[i % element.colors.length] ?? '#1D4ED8';
          return (
            <Group key={i}>
              <Rect
                x={bx}
                y={padT + plotH - bh}
                width={bw}
                height={bh}
                fill={color}
                cornerRadius={[4, 4, 0, 0]}
              />
              <Text
                x={padL + i * slot}
                y={padT + plotH + 8}
                width={slot}
                align="center"
                text={(labels[i] ?? '').slice(0, 11)}
                fontSize={11}
                fill="#475569"
              />
            </Group>
          );
        })}

      {element.chartKind === 'line' && (
        <>
          <Line
            points={linePoints}
            stroke={element.colors[0] ?? '#1D4ED8'}
            strokeWidth={3}
            lineCap="round"
            lineJoin="round"
          />
          {vals.map((v, i) => {
            const px = padL + ((i + 0.5) / vals.length) * plotW;
            const py = padT + plotH - (Math.max(0, v) / maxVal) * plotH;
            const slot = plotW / vals.length;
            return (
              <Group key={i}>
                <Circle
                  x={px}
                  y={py}
                  radius={4.5}
                  fill="#FFFFFF"
                  stroke={element.colors[0] ?? '#1D4ED8'}
                  strokeWidth={2.5}
                />
                <Text
                  x={padL + i * slot}
                  y={padT + plotH + 8}
                  width={slot}
                  align="center"
                  text={(labels[i] ?? '').slice(0, 11)}
                  fontSize={11}
                  fill="#475569"
                />
              </Group>
            );
          })}
        </>
      )}

      {element.chartKind === 'pie' && (
        <>
          {vals.map((v, i) => {
            const angle = (Math.max(0, v) / totalVal) * 360;
            const rot = pieStart;
            pieStart += angle;
            const radius = Math.min(plotW * 0.34, plotH * 0.46);
            const cx = padL + radius + 16;
            const cy = padT + plotH / 2;
            const color = element.colors[i % element.colors.length] ?? '#1D4ED8';
            return (
              <Group key={i}>
                <Wedge
                  x={cx}
                  y={cy}
                  radius={radius}
                  angle={angle}
                  rotation={rot}
                  fill={color}
                  stroke="#FFFFFF"
                  strokeWidth={1.5}
                />
                <Rect
                  x={cx + radius + 24}
                  y={padT + 14 + i * 24}
                  width={10}
                  height={10}
                  fill={color}
                  cornerRadius={2}
                />
                <Text
                  x={cx + radius + 40}
                  y={padT + 13 + i * 24}
                  width={Math.max(40, w - (cx + radius + 52))}
                  text={`${(labels[i] ?? '').slice(0, 12)} (${v})`}
                  fontSize={11}
                  fill="#334155"
                />
              </Group>
            );
          })}
        </>
      )}
    </Group>
  );
};

export const SlidesCanvas = forwardRef<SlidesCanvasHandle, SlidesCanvasProps>(
  (
    {
      slide,
      slideIndex,
      slideCount,
      themeId,
      aspectRatio = '16:9',
      master,
      selectedElementId,
      zoom,
      readOnly = false,
      onSelectElement,
      onUpdateElement,
    },
    ref,
  ) => {
    const containerRef = useRef<HTMLDivElement | null>(null);
    const stageRef = useRef<Konva.Stage | null>(null);
    const trRef = useRef<Konva.Transformer | null>(null);
    const nodeRefs = useRef<Record<string, Konva.Node | null>>({});
    const [fitScale, setFitScale] = useState(0.8);
    const [editingTextId, setEditingTextId] = useState<string | null>(null);
    const [guideX, setGuideX] = useState<number | null>(null);
    const [guideY, setGuideY] = useState<number | null>(null);

    const theme = useMemo(() => getSlideTheme(themeId), [themeId]);
    const dims = useMemo(() => getSlideDimensions(aspectRatio), [aspectRatio]);
    const slideW = dims.width;
    const slideH = dims.height;

    // Auto-fit stage scale to container bounds
    useEffect(() => {
      const el = containerRef.current;
      if (!el) return;
      const updateScale = () => {
        const rect = el.getBoundingClientRect();
        const availW = Math.max(260, rect.width - 36);
        const availH = Math.max(180, rect.height - 36);
        const s = Math.min(availW / slideW, availH / slideH, 1.25);
        setFitScale(Math.max(0.28, s));
      };
      updateScale();
      const ro = new ResizeObserver(updateScale);
      ro.observe(el);
      return () => ro.disconnect();
    }, [slideW, slideH]);

    const effectiveScale = Math.max(0.25, Math.min(2, fitScale * (zoom / 100)));

    // Attach Konva Transformer to selected node (edit mode only)
    useEffect(() => {
      const tr = trRef.current;
      if (!tr) return;
      if (readOnly || !selectedElementId || editingTextId === selectedElementId) {
        tr.nodes([]);
        tr.getLayer()?.batchDraw();
        return;
      }
      const node = nodeRefs.current[selectedElementId];
      if (node) {
        tr.nodes([node]);
        tr.getLayer()?.batchDraw();
      } else {
        tr.nodes([]);
        tr.getLayer()?.batchDraw();
      }
    }, [readOnly, selectedElementId, editingTextId, slide.elements]);

    // Close text editor when switching slides or entering read-only View mode
    useEffect(() => {
      setEditingTextId(null);
    }, [slide.id, readOnly]);

    useImperativeHandle(ref, () => ({
      exportPngDataUrl: () => {
        const stage = stageRef.current;
        if (!stage) return null;
        const tr = trRef.current;
        const prevNodes = tr ? tr.nodes() : [];
        if (tr) {
          tr.nodes([]);
          tr.getLayer()?.draw();
        }
        const dataUrl = stage.toDataURL({
          pixelRatio: Math.max(2, Math.round(2 / effectiveScale)),
          mimeType: 'image/png',
        });
        if (tr && prevNodes.length > 0) {
          tr.nodes(prevNodes);
          tr.getLayer()?.draw();
        }
        return dataUrl;
      },
    }));

    const handleSnapDragMove = (el: SlideElement, e: Konva.KonvaEventObject<DragEvent>) => {
      if (readOnly) return;
      const node = e.target;
      let nx = node.x();
      let ny = node.y();
      const cx = nx + el.width / 2;
      const cy = ny + el.height / 2;
      if (Math.abs(cx - slideW / 2) < 8) {
        nx = slideW / 2 - el.width / 2;
        node.x(nx);
        setGuideX(slideW / 2);
      } else {
        setGuideX(null);
      }
      if (Math.abs(cy - slideH / 2) < 8) {
        ny = slideH / 2 - el.height / 2;
        node.y(ny);
        setGuideY(slideH / 2);
      } else {
        setGuideY(null);
      }
    };

    const handleDragEnd = (id: string, x: number, y: number) => {
      setGuideX(null);
      setGuideY(null);
      if (readOnly) return;
      onUpdateElement(id, {x: Math.round(x), y: Math.round(y)});
    };

    const handleTransformEnd = (el: SlideElement, node: Konva.Node) => {
      if (readOnly) return;
      const scaleX = node.scaleX();
      const scaleY = node.scaleY();
      node.scaleX(1);
      node.scaleY(1);
      const nextW = Math.max(28, Math.round(el.width * scaleX));
      const nextH = Math.max(18, Math.round(el.height * scaleY));
      onUpdateElement(el.id, {
        x: Math.round(node.x()),
        y: Math.round(node.y()),
        width: nextW,
        height: nextH,
        rotation: Math.round(node.rotation()),
      });
    };

    const editingElement = useMemo(
      () =>
        readOnly
          ? null
          : (slide.elements.find(
              (e): e is TextElement => e.type === 'text' && e.id === editingTextId,
            ) ?? null),
      [readOnly, slide.elements, editingTextId],
    );

    const renderShapeElement = (el: ShapeElement) => {
      const common = {
        ref: (n: Konva.Node | null) => {
          nodeRefs.current[el.id] = n;
        },
        id: el.id,
        x: el.x,
        y: el.y,
        width: el.width,
        height: el.height,
        rotation: el.rotation,
        opacity: el.opacity ?? 1,
        draggable: !readOnly,
        onClick: readOnly ? undefined : () => onSelectElement(el.id),
        onTap: readOnly ? undefined : () => onSelectElement(el.id),
        onDragMove: (e: Konva.KonvaEventObject<DragEvent>) => handleSnapDragMove(el, e),
        onDragEnd: (e: Konva.KonvaEventObject<DragEvent>) =>
          handleDragEnd(el.id, e.target.x(), e.target.y()),
        onTransformEnd: (e: Konva.KonvaEventObject<Event>) => handleTransformEnd(el, e.target),
      };

      const strokeColor = el.stroke && el.stroke !== 'transparent' ? el.stroke : undefined;

      return (
        <Group key={el.id} {...common}>
          {el.shapeKind === 'rect' && (
            <Rect
              width={el.width}
              height={el.height}
              fill={el.fill}
              stroke={strokeColor}
              strokeWidth={el.strokeWidth}
            />
          )}
          {el.shapeKind === 'rounded-rect' && (
            <Rect
              width={el.width}
              height={el.height}
              fill={el.fill}
              stroke={strokeColor}
              strokeWidth={el.strokeWidth}
              cornerRadius={el.cornerRadius ?? 12}
            />
          )}
          {el.shapeKind === 'ellipse' && (
            <Ellipse
              x={el.width / 2}
              y={el.height / 2}
              radiusX={Math.max(4, el.width / 2)}
              radiusY={Math.max(4, el.height / 2)}
              fill={el.fill}
              stroke={strokeColor}
              strokeWidth={el.strokeWidth}
            />
          )}
          {el.shapeKind === 'triangle' && (
            <RegularPolygon
              x={el.width / 2}
              y={el.height * 0.58}
              sides={3}
              radius={Math.min(el.width, el.height) * 0.52}
              fill={el.fill}
              stroke={strokeColor}
              strokeWidth={el.strokeWidth}
            />
          )}
          {el.shapeKind === 'diamond' && (
            <RegularPolygon
              x={el.width / 2}
              y={el.height / 2}
              sides={4}
              radius={Math.min(el.width, el.height) * 0.5}
              fill={el.fill}
              stroke={strokeColor}
              strokeWidth={el.strokeWidth}
            />
          )}
          {el.shapeKind === 'arrow-right' && (
            <Arrow
              points={[8, el.height / 2, Math.max(24, el.width - 12), el.height / 2]}
              pointerLength={Math.min(28, el.width * 0.25)}
              pointerWidth={Math.min(34, el.height * 0.5)}
              fill={el.fill}
              stroke={el.fill}
              strokeWidth={Math.max(8, Math.min(22, el.height * 0.26))}
            />
          )}
          {el.shapeKind === 'line' && (
            <Line
              points={[0, el.height / 2, el.width, el.height / 2]}
              stroke={el.fill || theme.primary}
              strokeWidth={Math.max(3, el.strokeWidth || 3)}
              lineCap="round"
            />
          )}
          {el.label && el.shapeKind !== 'line' && (
            <Text
              width={el.width}
              height={el.height}
              text={el.label}
              fontSize={15}
              fontStyle="bold"
              fontFamily="Plus Jakarta Sans, sans-serif"
              fill={el.labelColor || '#FFFFFF'}
              align="center"
              verticalAlign="middle"
              padding={8}
            />
          )}
        </Group>
      );
    };

    return (
      <div
        ref={containerRef}
        data-testid="slides-canvas-stage"
        className="relative flex min-h-0 flex-1 items-center justify-center overflow-auto bg-slate-200/80 p-4 dark:bg-slate-950"
        onMouseDown={(e) => {
          if (e.target === e.currentTarget) {
            onSelectElement(null);
            setEditingTextId(null);
          }
        }}
      >
        <div
          data-testid="slides-canvas-frame"
          data-aspect-ratio={dims.aspectRatio}
          data-slide-width={slideW}
          data-slide-height={slideH}
          className="relative shrink-0 shadow-lg ring-1 ring-slate-300 dark:ring-slate-700"
          style={{
            width: Math.round(slideW * effectiveScale),
            height: Math.round(slideH * effectiveScale),
          }}
        >
          <Stage
            ref={stageRef}
            width={Math.round(slideW * effectiveScale)}
            height={Math.round(slideH * effectiveScale)}
            scaleX={effectiveScale}
            scaleY={effectiveScale}
            onMouseDown={(e) => {
              const clickedEmpty =
                e.target === e.target.getStage() || e.target.id() === 'slide-bg-rect';
              if (clickedEmpty) {
                onSelectElement(null);
                setEditingTextId(null);
              }
            }}
          >
            <Layer>
              {/* Slide Background */}
              <Rect
                id="slide-bg-rect"
                x={0}
                y={0}
                width={slideW}
                height={slideH}
                fill={slide.background || '#FFFFFF'}
              />

              {/* Slide Master Accent Bar */}
              {master.showAccentBar && (
                <Rect
                  x={0}
                  y={0}
                  width={slideW}
                  height={6}
                  fill={theme.primary}
                  listening={false}
                />
              )}

              {/* Elements */}
              {slide.elements.map((el) => {
                if (el.type === 'shape') {
                  return renderShapeElement(el);
                }

                if (el.type === 'text') {
                  const isEditing = editingTextId === el.id;
                  const fontStyle =
                    [el.bold ? 'bold' : '', el.italic ? 'italic' : '']
                      .filter(Boolean)
                      .join(' ') || 'normal';
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
                  const fontStack = getSlideFontCssStack(el.fontFamily);

                  return (
                    <Group
                      key={el.id}
                      ref={(n) => {
                        nodeRefs.current[el.id] = n;
                      }}
                      id={el.id}
                      x={el.x}
                      y={el.y}
                      width={el.width}
                      height={el.height}
                      rotation={el.rotation}
                      opacity={isEditing ? 0 : (el.opacity ?? 1)}
                      draggable={!readOnly && !isEditing}
                      onClick={readOnly ? undefined : () => onSelectElement(el.id)}
                      onTap={readOnly ? undefined : () => onSelectElement(el.id)}
                      onDblClick={
                        readOnly
                          ? undefined
                          : () => {
                              onSelectElement(el.id);
                              setEditingTextId(el.id);
                            }
                      }
                      onDblTap={
                        readOnly
                          ? undefined
                          : () => {
                              onSelectElement(el.id);
                              setEditingTextId(el.id);
                            }
                      }
                      onDragMove={(e) => handleSnapDragMove(el, e)}
                      onDragEnd={(e) => handleDragEnd(el.id, e.target.x(), e.target.y())}
                      onTransformEnd={(e) => handleTransformEnd(el, e.target)}
                    >
                      <Rect
                        width={el.width}
                        height={el.height}
                        fill={el.fill && el.fill !== 'transparent' ? el.fill : 'transparent'}
                        stroke={fit.stillOverflows ? '#D97706' : undefined}
                        strokeWidth={fit.stillOverflows ? 1.5 : 0}
                        dash={fit.stillOverflows ? [6, 4] : undefined}
                        cornerRadius={4}
                      />
                      <Text
                        width={el.width}
                        height={el.height}
                        text={formatListDisplayText(el.text, el.listStyle)}
                        fontFamily={fontStack}
                        fontSize={fit.fittedFontSize}
                        fontStyle={fontStyle}
                        textDecoration={el.underline ? 'underline' : ''}
                        align={el.align ?? 'left'}
                        fill={el.color}
                        lineHeight={fit.fittedLineHeight}
                        padding={el.padding ?? 4}
                        wrap="word"
                      />
                    </Group>
                  );
                }

                if (el.type === 'image') {
                  return (
                    <LoadedKonvaImage
                      key={el.id}
                      element={el}
                      readOnly={readOnly}
                      setRef={(n) => {
                        nodeRefs.current[el.id] = n;
                      }}
                      onSelect={() => onSelectElement(el.id)}
                      onDragEnd={(x, y) => handleDragEnd(el.id, x, y)}
                      onTransformEnd={(n) => handleTransformEnd(el, n)}
                    />
                  );
                }

                if (el.type === 'table') {
                  return (
                    <KonvaTableNode
                      key={el.id}
                      element={el}
                      readOnly={readOnly}
                      setRef={(n) => {
                        nodeRefs.current[el.id] = n;
                      }}
                      onSelect={() => onSelectElement(el.id)}
                      onDragEnd={(x, y) => handleDragEnd(el.id, x, y)}
                      onTransformEnd={(n) => handleTransformEnd(el, n)}
                    />
                  );
                }

                return (
                  <KonvaChartNode
                    key={el.id}
                    element={el}
                    readOnly={readOnly}
                    setRef={(n) => {
                      nodeRefs.current[el.id] = n;
                    }}
                    onSelect={() => onSelectElement(el.id)}
                    onDragEnd={(x, y) => handleDragEnd(el.id, x, y)}
                    onTransformEnd={(n) => handleTransformEnd(el, n)}
                  />
                );
              })}

              {/* Slide Master Footer & Slide Number */}
              {master.footerText.trim() && (
                <Text
                  x={42}
                  y={slideH - 28}
                  width={600}
                  text={master.footerText}
                  fontSize={11}
                  fontFamily="DM Sans, sans-serif"
                  fill={theme.mutedColor}
                  listening={false}
                />
              )}
              {master.showSlideNumbers && (
                <Text
                  x={slideW - 96}
                  y={slideH - 28}
                  width={56}
                  align="right"
                  text={`${slideIndex + 1} / ${slideCount}`}
                  fontSize={11}
                  fontFamily="DM Sans, sans-serif"
                  fill={theme.mutedColor}
                  listening={false}
                />
              )}

              {/* Snap center alignment guides */}
              {guideX !== null && (
                <Line
                  points={[guideX, 0, guideX, slideH]}
                  stroke="#EA580C"
                  strokeWidth={1}
                  dash={[4, 4]}
                  listening={false}
                />
              )}
              {guideY !== null && (
                <Line
                  points={[0, guideY, slideW, guideY]}
                  stroke="#EA580C"
                  strokeWidth={1}
                  dash={[4, 4]}
                  listening={false}
                />
              )}

              {!readOnly && (
                <Transformer
                  ref={trRef}
                  rotateEnabled
                  borderStroke="#EA580C"
                  anchorStroke="#EA580C"
                  anchorFill="#FFFFFF"
                  anchorSize={8}
                  boundBoxFunc={(oldBox, newBox) => {
                    if (newBox.width < 24 || newBox.height < 16) return oldBox;
                    return newBox;
                  }}
                />
              )}
            </Layer>
          </Stage>

          {/* Exact-position Inline Text Editor Overlay */}
          {editingElement && (
            <textarea
              aria-label="Edit slide text"
              autoFocus
              value={editingElement.text}
              onChange={(e) => onUpdateElement(editingElement.id, {text: e.target.value})}
              onBlur={() => setEditingTextId(null)}
              onKeyDown={(e) => {
                if (e.key === 'Escape') {
                  e.stopPropagation();
                  setEditingTextId(null);
                }
              }}
              style={{
                position: 'absolute',
                left: editingElement.x * effectiveScale,
                top: editingElement.y * effectiveScale,
                width: Math.max(120, editingElement.width * effectiveScale),
                height: Math.max(48, editingElement.height * effectiveScale),
                fontFamily: getSlideFontCssStack(editingElement.fontFamily),
                fontSize: Math.max(11, editingElement.fontSize * effectiveScale),
                fontWeight: editingElement.bold ? 700 : 400,
                fontStyle: editingElement.italic ? 'italic' : 'normal',
                textAlign: editingElement.align ?? 'left',
                color: editingElement.color,
                lineHeight: editingElement.lineHeight ?? 1.35,
              }}
              className="resize-none rounded border-2 border-orange-500 bg-white/95 p-1 outline-none dark:bg-slate-900/95"
            />
          )}
        </div>
      </div>
    );
  },
);

SlidesCanvas.displayName = 'SlidesCanvas';

/**
 * Lightweight, zero-overhead visual slide renderer used for Filmstrip thumbnails
 * and Full-Screen Presenter Mode (smooth CSS transitions & crisp text legibility).
 */
export const SlideSvgRenderer: React.FC<{
  slide: SlideItem;
  slideIndex: number;
  slideCount: number;
  themeId: string;
  aspectRatio?: SlideAspectRatio;
  master: SlideMasterSettings;
  className?: string;
}> = ({slide, slideIndex, slideCount, themeId, aspectRatio = '16:9', master, className = ''}) => {
  const theme = getSlideTheme(themeId);
  const dims = getSlideDimensions(aspectRatio);
  const slideW = dims.width;
  const slideH = dims.height;

  return (
    <svg
      viewBox={`0 0 ${slideW} ${slideH}`}
      className={`block h-full w-full select-none ${className}`}
      role="img"
      aria-label={`Slide ${slideIndex + 1}`}
    >
      <rect x={0} y={0} width={slideW} height={slideH} fill={slide.background || '#FFFFFF'} />
      {master.showAccentBar && (
        <rect x={0} y={0} width={slideW} height={6} fill={theme.primary} />
      )}

      {slide.elements.map((el) => {
        if (el.type === 'shape') {
          return (
            <g key={el.id} transform={`translate(${el.x}, ${el.y}) rotate(${el.rotation || 0})`}>
              {el.shapeKind === 'ellipse' ? (
                <ellipse
                  cx={el.width / 2}
                  cy={el.height / 2}
                  rx={el.width / 2}
                  ry={el.height / 2}
                  fill={el.fill}
                />
              ) : el.shapeKind === 'triangle' ? (
                <polygon
                  points={`${el.width / 2},0 ${el.width},${el.height} 0,${el.height}`}
                  fill={el.fill}
                />
              ) : el.shapeKind === 'diamond' ? (
                <polygon
                  points={`${el.width / 2},0 ${el.width},${el.height / 2} ${el.width / 2},${el.height} 0,${el.height / 2}`}
                  fill={el.fill}
                />
              ) : el.shapeKind === 'line' ? (
                <line
                  x1={0}
                  y1={el.height / 2}
                  x2={el.width}
                  y2={el.height / 2}
                  stroke={el.fill || theme.primary}
                  strokeWidth={Math.max(3, el.strokeWidth || 3)}
                />
              ) : (
                <rect
                  width={el.width}
                  height={el.height}
                  rx={el.shapeKind === 'rounded-rect' ? (el.cornerRadius ?? 12) : 0}
                  fill={el.fill}
                />
              )}
              {el.label && el.shapeKind !== 'line' && (
                <text
                  x={el.width / 2}
                  y={el.height / 2 + 5}
                  textAnchor="middle"
                  fill={el.labelColor || '#FFFFFF'}
                  fontSize={15}
                  fontWeight={700}
                >
                  {el.label}
                </text>
              )}
            </g>
          );
        }

        if (el.type === 'text') {
          const formatted = formatListDisplayText(el.text, el.listStyle);
          const lines = formatted.split('\n');
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
          const lineStep = fit.fittedFontSize * fit.fittedLineHeight;
          const anchorX =
            el.align === 'center' ? el.width / 2 : el.align === 'right' ? el.width - 4 : 4;
          const textAnchor =
            el.align === 'center' ? 'middle' : el.align === 'right' ? 'end' : 'start';
          return (
            <g key={el.id} transform={`translate(${el.x}, ${el.y}) rotate(${el.rotation || 0})`}>
              {el.fill && el.fill !== 'transparent' && (
                <rect width={el.width} height={el.height} fill={el.fill} rx={4} />
              )}
              <foreignObject x={0} y={0} width={el.width} height={el.height}>
                <div
                  style={{
                    width: '100%',
                    height: '100%',
                    boxSizing: 'border-box',
                    fontFamily: getSlideFontCssStack(el.fontFamily),
                    fontSize: `${fit.fittedFontSize}px`,
                    fontWeight: el.bold ? 700 : 400,
                    fontStyle: el.italic ? 'italic' : 'normal',
                    textDecoration: el.underline ? 'underline' : 'none',
                    textAlign: el.align ?? 'left',
                    color: el.color,
                    lineHeight: fit.fittedLineHeight,
                    whiteSpace: 'pre-wrap',
                    overflowWrap: 'break-word',
                    overflow: 'hidden',
                    padding: `${el.padding ?? 4}px`,
                  }}
                >
                  {lines.length > 0 ? formatted : ''}
                </div>
              </foreignObject>
              {/* Accessible SVG text fallback if foreignObject is ignored */}
              <title>{`${anchorX}:${textAnchor}:${lineStep}:${el.text}`}</title>
            </g>
          );
        }

        if (el.type === 'image') {
          return (
            <image
              key={el.id}
              href={el.src}
              x={el.x}
              y={el.y}
              width={el.width}
              height={el.height}
              preserveAspectRatio="xMidYMid slice"
            />
          );
        }

        if (el.type === 'table') {
          const rowCount = Math.max(1, el.rows.length);
          const colCount = Math.max(1, el.rows[0]?.length ?? 1);
          const rowH = el.height / rowCount;
          const colW = el.width / colCount;
          return (
            <g key={el.id} transform={`translate(${el.x}, ${el.y})`}>
              {el.rows.map((row, rIdx) =>
                row.map((cell, cIdx) => (
                  <g key={`${rIdx}-${cIdx}`} transform={`translate(${cIdx * colW}, ${rIdx * rowH})`}>
                    <rect
                      width={colW}
                      height={rowH}
                      fill={rIdx === 0 ? el.headerBg : el.cellBg}
                      stroke={el.borderColor}
                      strokeWidth={1}
                    />
                    <text
                      x={10}
                      y={rowH / 2 + 5}
                      fontSize={el.fontSize}
                      fontWeight={rIdx === 0 ? 700 : 400}
                      fill={rIdx === 0 ? el.headerColor : el.cellColor}
                    >
                      {cell.slice(0, 26)}
                    </text>
                  </g>
                )),
              )}
            </g>
          );
        }

        // Chart element
        const vals = el.values.length > 0 ? el.values : [10, 20, 30];
        const maxVal = Math.max(1, ...vals);
        const plotW = Math.max(40, el.width - 56);
        const plotH = Math.max(40, el.height - 76);
        return (
          <g key={el.id} transform={`translate(${el.x}, ${el.y})`}>
            <rect
              width={el.width}
              height={el.height}
              rx={10}
              fill="#FFFFFF"
              stroke="#E2E8F0"
              strokeWidth={1.2}
            />
            <text x={16} y={26} fontSize={14} fontWeight={700} fill="#0F172A">
              {el.title}
            </text>
            {el.chartKind === 'bar' &&
              vals.map((v, i) => {
                const slot = plotW / vals.length;
                const bw = Math.max(12, slot * 0.56);
                const bh = Math.max(4, (Math.max(0, v) / maxVal) * plotH);
                return (
                  <g key={i}>
                    <rect
                      x={32 + i * slot + (slot - bw) / 2}
                      y={42 + plotH - bh}
                      width={bw}
                      height={bh}
                      rx={3}
                      fill={el.colors[i % el.colors.length] ?? theme.primary}
                    />
                    <text
                      x={32 + i * slot + slot / 2}
                      y={42 + plotH + 18}
                      textAnchor="middle"
                      fontSize={11}
                      fill="#475569"
                    >
                      {(el.labels[i] ?? '').slice(0, 11)}
                    </text>
                  </g>
                );
              })}
            {el.chartKind === 'line' && (
              <polyline
                fill="none"
                stroke={el.colors[0] ?? theme.primary}
                strokeWidth={3}
                points={vals
                  .map((v, i) => {
                    const px = 32 + ((i + 0.5) / vals.length) * plotW;
                    const py = 42 + plotH - (Math.max(0, v) / maxVal) * plotH;
                    return `${px},${py}`;
                  })
                  .join(' ')}
              />
            )}
            {el.chartKind === 'pie' && (
              <circle
                cx={el.width / 2}
                cy={el.height / 2 + 12}
                r={Math.min(plotW, plotH) * 0.42}
                fill={el.colors[0] ?? theme.primary}
              />
            )}
          </g>
        );
      })}

      {master.footerText.trim() && (
        <text x={42} y={slideH - 16} fontSize={11} fill={theme.mutedColor}>
          {master.footerText}
        </text>
      )}
      {master.showSlideNumbers && (
        <text
          x={slideW - 42}
          y={slideH - 16}
          textAnchor="end"
          fontSize={11}
          fill={theme.mutedColor}
        >
          {slideIndex + 1} / {slideCount}
        </text>
      )}
    </svg>
  );
};
