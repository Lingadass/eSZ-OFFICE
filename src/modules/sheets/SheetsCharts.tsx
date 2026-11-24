import React, {useMemo} from 'react';
import {BarChart3, Trash2, X} from 'lucide-react';
import {parseRangeRef} from './formulaEngine';
import {
  cellKey,
  CHART_COLORS,
  type ChartKind,
  type ChartSpec,
  type EvaluatedCell,
  type SheetData,
} from './sheetsTypes';

interface SheetsChartsPanelProps {
  sheet: SheetData;
  evalMap: Record<string, EvaluatedCell>;
  onChangeChartKind: (chartId: string, kind: ChartKind) => void;
  onDeleteChart: (chartId: string) => void;
  onClose: () => void;
  readOnly?: boolean;
}

interface ParsedChartSeries {
  labels: string[];
  series: Array<{name: string; values: number[]; color: string}>;
}

function extractChartData(
  sheet: SheetData,
  evalMap: Record<string, EvaluatedCell>,
  rangeStr: string,
): ParsedChartSeries {
  const parsed = parseRangeRef(rangeStr);
  if (!parsed) return {labels: [], series: []};
  const {startRow, endRow, startCol, endCol} = parsed;

  // Detect if first row looks like text headers
  let hasHeader = false;
  if (endRow > startRow) {
    const firstSeriesTop = evalMap[cellKey(startRow, Math.min(endCol, startCol + 1))]?.value;
    if (typeof firstSeriesTop === 'string') hasHeader = true;
  }

  const dataStartRow = hasHeader ? startRow + 1 : startRow;
  const labels: string[] = [];
  for (let r = dataStartRow; r <= endRow; r++) {
    const lbl = evalMap[cellKey(r, startCol)]?.display || sheet.cells[cellKey(r, startCol)]?.raw || `Row ${r + 1}`;
    labels.push(lbl);
  }

  const seriesStartCol = endCol > startCol ? startCol + 1 : startCol;
  const series: ParsedChartSeries['series'] = [];
  for (let c = seriesStartCol; c <= endCol; c++) {
    const name = hasHeader
      ? evalMap[cellKey(startRow, c)]?.display || `Series ${c - seriesStartCol + 1}`
      : `Series ${c - seriesStartCol + 1}`;
    const values: number[] = [];
    for (let r = dataStartRow; r <= endRow; r++) {
      const val = evalMap[cellKey(r, c)]?.value;
      values.push(typeof val === 'number' && Number.isFinite(val) ? val : 0);
    }
    series.push({
      name,
      values,
      color: CHART_COLORS[(c - seriesStartCol) % CHART_COLORS.length],
    });
  }

  return {labels, series};
}

const SingleChartCard: React.FC<{
  chart: ChartSpec;
  sheet: SheetData;
  evalMap: Record<string, EvaluatedCell>;
  onChangeKind: (kind: ChartKind) => void;
  onDelete: () => void;
  readOnly?: boolean;
}> = ({chart, sheet, evalMap, onChangeKind, onDelete, readOnly = false}) => {
  const data = useMemo(
    () => extractChartData(sheet, evalMap, chart.range),
    [sheet, evalMap, chart.range],
  );

  const w = 310;
  const h = 190;
  const padL = 38;
  const padR = 14;
  const padT = 16;
  const padB = 34;
  const plotW = w - padL - padR;
  const plotH = h - padT - padB;

  const allVals = data.series.flatMap((s) => s.values);
  const maxVal = Math.max(1, ...allVals.map((v) => Math.max(0, v)));

  const renderSvg = () => {
    if (data.labels.length === 0 || data.series.length === 0) {
      return (
        <div className="flex h-44 items-center justify-center text-xs text-slate-400">
          No numeric data in range {chart.range}
        </div>
      );
    }

    if (chart.kind === 'pie') {
      const vals = data.series[0].values.map((v) => Math.max(0, v));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = 100;
      const cy = 94;
      const r = 68;
      let startAngle = -Math.PI / 2;

      return (
        <svg viewBox={`0 0 ${w} ${h}`} className="h-48 w-full" role="img" aria-label={chart.title}>
          {vals.map((val, i) => {
            const frac = val / total;
            const angle = frac * Math.PI * 2;
            const endAngle = startAngle + angle;
            const x1 = cx + r * Math.cos(startAngle);
            const y1 = cy + r * Math.sin(startAngle);
            const x2 = cx + r * Math.cos(endAngle);
            const y2 = cy + r * Math.sin(endAngle);
            const largeArc = angle > Math.PI ? 1 : 0;
            const d =
              frac >= 0.999
                ? `M ${cx} ${cy - r} A ${r} ${r} 0 1 1 ${cx - 0.01} ${cy - r} Z`
                : `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;
            startAngle = endAngle;
            const color = CHART_COLORS[i % CHART_COLORS.length];
            return (
              <path key={i} d={d} fill={color} stroke="#ffffff" strokeWidth={1.5}>
                <title>
                  {data.labels[i]}: {val.toLocaleString('en-US')} ({Math.round(frac * 100)}%)
                </title>
              </path>
            );
          })}
          {/* Pie Legend */}
          {data.labels.slice(0, 6).map((lbl, i) => (
            <g key={i} transform={`translate(184, ${26 + i * 22})`}>
              <rect width={10} height={10} rx={2} fill={CHART_COLORS[i % CHART_COLORS.length]} />
              <text x={15} y={9} fontSize={10} fill="currentColor" className="text-slate-700 dark:text-slate-300">
                {lbl.slice(0, 14)}
              </text>
            </g>
          ))}
        </svg>
      );
    }

    const groupW = plotW / Math.max(1, data.labels.length);

    return (
      <svg viewBox={`0 0 ${w} ${h}`} className="h-48 w-full" role="img" aria-label={chart.title}>
        {/* Horizontal gridlines */}
        {[0, 0.25, 0.5, 0.75, 1].map((tick) => {
          const y = padT + plotH * (1 - tick);
          const v = Math.round(maxVal * tick);
          return (
            <g key={tick}>
              <line
                x1={padL}
                y1={y}
                x2={w - padR}
                y2={y}
                stroke="currentColor"
                strokeOpacity={0.12}
                strokeDasharray={tick === 0 ? undefined : '2 2'}
              />
              <text
                x={padL - 5}
                y={y + 3}
                textAnchor="end"
                fontSize={9}
                fill="currentColor"
                opacity={0.6}
              >
                {v >= 1000 ? `${(v / 1000).toFixed(1)}k` : v}
              </text>
            </g>
          );
        })}

        {/* Bar chart */}
        {chart.kind === 'bar' &&
          data.labels.map((lbl, i) => {
            const barCount = data.series.length;
            const innerW = Math.max(6, (groupW * 0.72) / barCount);
            const startX = padL + i * groupW + groupW * 0.14;
            return (
              <g key={i}>
                {data.series.map((ser, sIdx) => {
                  const val = Math.max(0, ser.values[i] ?? 0);
                  const bh = (val / maxVal) * plotH;
                  return (
                    <rect
                      key={sIdx}
                      x={startX + sIdx * innerW}
                      y={padT + plotH - bh}
                      width={Math.max(3, innerW - 2)}
                      height={bh}
                      rx={2}
                      fill={ser.color}
                    >
                      <title>
                        {lbl} — {ser.name}: {val.toLocaleString('en-US')}
                      </title>
                    </rect>
                  );
                })}
                <text
                  x={padL + i * groupW + groupW / 2}
                  y={h - 12}
                  textAnchor="middle"
                  fontSize={9}
                  fill="currentColor"
                  opacity={0.75}
                >
                  {lbl.length > 12 ? `${lbl.slice(0, 11)}…` : lbl}
                </text>
              </g>
            );
          })}

        {/* Line chart */}
        {chart.kind === 'line' && (
          <>
            {data.series.map((ser, sIdx) => {
              const pts = ser.values
                .map((val, i) => {
                  const x = padL + i * groupW + groupW / 2;
                  const y = padT + plotH - (Math.max(0, val) / maxVal) * plotH;
                  return `${x},${y}`;
                })
                .join(' ');
              return (
                <g key={sIdx}>
                  <polyline
                    fill="none"
                    stroke={ser.color}
                    strokeWidth={2.2}
                    strokeLinejoin="round"
                    points={pts}
                  />
                  {ser.values.map((val, i) => {
                    const x = padL + i * groupW + groupW / 2;
                    const y = padT + plotH - (Math.max(0, val) / maxVal) * plotH;
                    return (
                      <circle key={i} cx={x} cy={y} r={3.2} fill={ser.color}>
                        <title>
                          {data.labels[i]} — {ser.name}: {val.toLocaleString('en-US')}
                        </title>
                      </circle>
                    );
                  })}
                </g>
              );
            })}
            {data.labels.map((lbl, i) => (
              <text
                key={i}
                x={padL + i * groupW + groupW / 2}
                y={h - 12}
                textAnchor="middle"
                fontSize={9}
                fill="currentColor"
                opacity={0.75}
              >
                {lbl.length > 12 ? `${lbl.slice(0, 11)}…` : lbl}
              </text>
            ))}
          </>
        )}

        {/* Scatter plot */}
        {chart.kind === 'scatter' && (
          <>
            {data.series.map((ser, sIdx) =>
              ser.values.map((val, i) => {
                const x = padL + ((i + 0.5) / Math.max(1, data.labels.length)) * plotW;
                const y = padT + plotH - (Math.max(0, val) / maxVal) * plotH;
                return (
                  <circle
                    key={`${sIdx}-${i}`}
                    cx={x}
                    cy={y}
                    r={4.2}
                    fill={ser.color}
                    fillOpacity={0.85}
                  >
                    <title>
                      {data.labels[i]} ({ser.name}): {val.toLocaleString('en-US')}
                    </title>
                  </circle>
                );
              }),
            )}
            {data.labels.map((lbl, i) => (
              <text
                key={i}
                x={padL + ((i + 0.5) / Math.max(1, data.labels.length)) * plotW}
                y={h - 12}
                textAnchor="middle"
                fontSize={9}
                fill="currentColor"
                opacity={0.75}
              >
                {lbl.length > 12 ? `${lbl.slice(0, 11)}…` : lbl}
              </text>
            ))}
          </>
        )}
      </svg>
    );
  };

  return (
    <div
      data-testid="sheet-chart-card"
      className="rounded-lg border border-slate-200 bg-white p-3 shadow-xs dark:border-slate-800 dark:bg-slate-950"
    >
      <div className="mb-2 flex items-center justify-between gap-2">
        <div className="min-w-0">
          <div className="truncate text-xs font-semibold text-slate-900 dark:text-slate-100">
            {chart.title}
          </div>
          <div className="font-mono text-[10px] text-slate-400">Range: {chart.range}</div>
        </div>
        {!readOnly && (
          <div className="flex shrink-0 items-center gap-1">
            <select
              aria-label={`Chart type for ${chart.title}`}
              value={chart.kind}
              onChange={(e) => onChangeKind(e.target.value as ChartKind)}
              className="h-6 rounded border border-slate-200 bg-slate-50 px-1 text-[11px] dark:border-slate-700 dark:bg-slate-900"
            >
              <option value="bar">Bar</option>
              <option value="line">Line</option>
              <option value="pie">Pie</option>
              <option value="scatter">Scatter</option>
            </select>
            <button
              type="button"
              onClick={onDelete}
              aria-label={`Delete chart ${chart.title}`}
              className="rounded p-1 text-slate-400 hover:bg-rose-50 hover:text-rose-600 dark:hover:bg-rose-950/50"
            >
              <Trash2 className="h-3.5 w-3.5" />
            </button>
          </div>
        )}
      </div>

      {renderSvg()}

      {chart.kind !== 'pie' && data.series.length > 0 && (
        <div className="mt-1.5 flex flex-wrap gap-2 border-t border-slate-100 pt-1.5 text-[10px] text-slate-600 dark:border-slate-800 dark:text-slate-400">
          {data.series.map((s) => (
            <span key={s.name} className="inline-flex items-center gap-1">
              <span
                className="inline-block h-2 w-2 rounded-xs"
                style={{backgroundColor: s.color}}
              />
              {s.name}
            </span>
          ))}
        </div>
      )}
    </div>
  );
};

export const SheetsChartsPanel: React.FC<SheetsChartsPanelProps> = ({
  sheet,
  evalMap,
  onChangeChartKind,
  onDeleteChart,
  onClose,
  readOnly = false,
}) => (
  <aside
    aria-label="Embedded sheet charts"
    className="flex w-full shrink-0 flex-col border-l border-slate-200 bg-slate-50 sm:w-[350px] dark:border-slate-800 dark:bg-slate-900"
  >
    <div className="flex items-center justify-between border-b border-slate-200 bg-white px-3.5 py-2.5 dark:border-slate-800 dark:bg-slate-900">
      <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-900 dark:text-slate-100">
        <BarChart3 className="h-3.5 w-3.5 text-emerald-700 dark:text-emerald-400" />
        Embedded Charts ({sheet.charts.length})
      </div>
      <button
        type="button"
        onClick={onClose}
        aria-label="Close charts panel"
        className="rounded p-1 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
      >
        <X className="h-4 w-4" />
      </button>
    </div>

    <div className="flex-1 space-y-3 overflow-y-auto p-3">
      {sheet.charts.map((chart) => (
        <SingleChartCard
          key={chart.id}
          chart={chart}
          sheet={sheet}
          evalMap={evalMap}
          readOnly={readOnly}
          onChangeKind={(k) => onChangeChartKind(chart.id, k)}
          onDelete={() => onDeleteChart(chart.id)}
        />
      ))}
    </div>
  </aside>
);
