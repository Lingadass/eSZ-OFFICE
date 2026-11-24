import React, {useEffect, useRef, useState} from 'react';
import type {PDFDocumentProxy} from 'pdfjs-dist';
import {
  ArrowDown,
  ArrowUp,
  Check,
  FilePlus,
  Layers,
  Minimize2,
  RotateCcw,
  RotateCw,
  Scissors,
  Trash2,
  X,
} from 'lucide-react';
import type {OfficeFileRecord} from '../../core/types';
import {formatBytes} from '../../core/storage/fileRules';
import {loadFileBlob} from '../../core/storage/fileRepository';
import {compressPdfWithQpdf} from './pdfEngine';
import {
  extractPdfPagesBytes,
  flattenPdfBytes,
  mergePdfBytes,
  parsePageRange,
  readPdfMetadata,
  reorderAndTransformPdfBytes,
  splitIntoChunkSpecs,
  updatePdfMetadata,
} from './pdfPure';
import type {PdfMetadataInfo} from './pdfTypes';

interface OrganizePageItem {
  key: string;
  sourceIndex: number | 'blank-a4' | 'blank-letter';
  label: string;
  extraRotationDeg: number;
  selected: boolean;
}

interface PdfOrganizeModalProps {
  open: boolean;
  onClose: () => void;
  fileName: string;
  pdfBytes: Uint8Array;
  pdfDoc: PDFDocumentProxy | null;
  vaultPdfFiles: OfficeFileRecord[];
  onApplyBytes: (nextBytes: Uint8Array, summary: string) => void;
  onCreateNewPdf: (name: string, bytes: Uint8Array) => Promise<void>;
  onToast: (msg: string, tone?: 'info' | 'success' | 'warning' | 'error') => void;
}

const MiniPageThumb: React.FC<{
  pdfDoc: PDFDocumentProxy | null;
  pageNumber: number | null;
  rotationDeg: number;
}> = ({pdfDoc, pageNumber, rotationDeg}) => {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);

  useEffect(() => {
    let cancelled = false;
    if (!pdfDoc || !pageNumber) return;

    const draw = async () => {
      try {
        const page = await pdfDoc.getPage(pageNumber);
        if (cancelled) return;
        const rot = (page.rotate + rotationDeg + 360) % 360;
        const unscaled = page.getViewport({scale: 1, rotation: rot});
        const scale = Math.min(110 / unscaled.width, 136 / unscaled.height);
        const viewport = page.getViewport({scale, rotation: rot});
        const canvas = canvasRef.current;
        if (!canvas || cancelled) return;
        canvas.width = Math.max(1, Math.floor(viewport.width));
        canvas.height = Math.max(1, Math.floor(viewport.height));
        const ctx = canvas.getContext('2d');
        if (!ctx) return;
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({canvasContext: ctx, viewport, canvas}).promise;
      } catch {
        // Ignore thumbnail cancel/render error
      }
    };
    void draw();
    return () => {
      cancelled = true;
    };
  }, [pdfDoc, pageNumber, rotationDeg]);

  if (!pageNumber) {
    return (
      <div className="flex h-32 w-24 items-center justify-center rounded border border-dashed border-slate-300 bg-white text-[11px] font-medium text-slate-400 dark:border-slate-700 dark:bg-slate-950">
        Blank Page
      </div>
    );
  }

  return (
    <canvas
      ref={canvasRef}
      className="max-h-32 max-w-[110px] rounded border border-slate-200 bg-white shadow-xs dark:border-slate-700"
    />
  );
};

export const PdfOrganizeModal: React.FC<PdfOrganizeModalProps> = ({
  open,
  onClose,
  fileName,
  pdfBytes,
  pdfDoc,
  vaultPdfFiles,
  onApplyBytes,
  onCreateNewPdf,
  onToast,
}) => {
  const [section, setSection] = useState<'pages' | 'merge' | 'split' | 'optimize'>('pages');
  const [busy, setBusy] = useState(false);

  // Pages state
  const [pages, setPages] = useState<OrganizePageItem[]>([]);
  const [dragIdx, setDragIdx] = useState<number | null>(null);

  // Merge state
  const [mergeSelectedIds, setMergeSelectedIds] = useState<string[]>([]);
  const [mergeLocalFiles, setMergeLocalFiles] = useState<File[]>([]);

  // Split state
  const [splitMode, setSplitMode] = useState<'chunks' | 'range'>('chunks');
  const [pagesPerChunk, setPagesPerChunk] = useState(1);
  const [rangeSpec, setRangeSpec] = useState('1-2');

  // Metadata state
  const [meta, setMeta] = useState<PdfMetadataInfo>({
    title: '',
    author: '',
    subject: '',
    keywords: '',
    creator: 'eSZ OFFICE',
  });

  const numPages = pdfDoc?.numPages ?? 0;

  useEffect(() => {
    if (!open) return;
    setPages(
      Array.from({length: numPages}, (_, i) => ({
        key: `orig_${i}_${Date.now()}`,
        sourceIndex: i,
        label: `Page ${i + 1}`,
        extraRotationDeg: 0,
        selected: false,
      })),
    );
    setMergeSelectedIds([]);
    setMergeLocalFiles([]);
    void readPdfMetadata(pdfBytes).then(setMeta).catch(() => undefined);
  }, [open, numPages, pdfBytes]);

  if (!open) return null;

  const selectedIndices = pages
    .map((p, idx) => (p.selected ? idx : -1))
    .filter((i) => i >= 0);

  const toggleSelect = (idx: number) => {
    setPages((prev) => prev.map((p, i) => (i === idx ? {...p, selected: !p.selected} : p)));
  };

  const rotateSelected = (delta: number) => {
    if (selectedIndices.length === 0) {
      onToast('Select one or more pages first.', 'info');
      return;
    }
    setPages((prev) =>
      prev.map((p) =>
        p.selected ? {...p, extraRotationDeg: (p.extraRotationDeg + delta + 360) % 360} : p,
      ),
    );
  };

  const deleteSelected = () => {
    if (selectedIndices.length === 0) return;
    if (selectedIndices.length >= pages.length) {
      onToast('A PDF must keep at least 1 page.', 'warning');
      return;
    }
    setPages((prev) => prev.filter((p) => !p.selected));
  };

  const insertBlank = (kind: 'blank-a4' | 'blank-letter') => {
    const newItem: OrganizePageItem = {
      key: `blank_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      sourceIndex: kind,
      label: kind === 'blank-a4' ? 'Blank A4' : 'Blank Letter',
      extraRotationDeg: 0,
      selected: false,
    };
    const lastSel = selectedIndices[selectedIndices.length - 1];
    if (lastSel !== undefined) {
      setPages((prev) => [
        ...prev.slice(0, lastSel + 1),
        newItem,
        ...prev.slice(lastSel + 1),
      ]);
    } else {
      setPages((prev) => [...prev, newItem]);
    }
  };

  const movePage = (from: number, to: number) => {
    if (to < 0 || to >= pages.length || from === to) return;
    setPages((prev) => {
      const copy = [...prev];
      const [moved] = copy.splice(from, 1);
      copy.splice(to, 0, moved);
      return copy;
    });
  };

  const handleApplyPageOrder = async () => {
    setBusy(true);
    try {
      const nextBytes = await reorderAndTransformPdfBytes(
        pdfBytes,
        pages.map((p) => ({
          sourceIndex: p.sourceIndex,
          extraRotationDeg: p.extraRotationDeg,
        })),
      );
      onApplyBytes(nextBytes, 'Updated page organization');
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Failed to update pages.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleExtractSelected = async () => {
    const origTargets = pages
      .filter((p) => p.selected && typeof p.sourceIndex === 'number')
      .map((p) => p.sourceIndex as number);
    if (origTargets.length === 0) {
      onToast('Select original pages to extract.', 'info');
      return;
    }
    setBusy(true);
    try {
      const extracted = await extractPdfPagesBytes(pdfBytes, origTargets);
      const base = fileName.replace(/\.pdf$/i, '');
      await onCreateNewPdf(`${base} (Extracted).pdf`, extracted);
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Extract failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleMerge = async (mode: 'append' | 'new') => {
    if (mergeSelectedIds.length === 0 && mergeLocalFiles.length === 0) {
      onToast('Choose at least one PDF to merge.', 'info');
      return;
    }
    setBusy(true);
    try {
      const buffers: Uint8Array[] = [pdfBytes];
      for (const id of mergeSelectedIds) {
        const blob = await loadFileBlob(id);
        if (blob) {
          buffers.push(new Uint8Array(await blob.arrayBuffer()));
        }
      }
      for (const f of mergeLocalFiles) {
        buffers.push(new Uint8Array(await f.arrayBuffer()));
      }
      const merged = await mergePdfBytes(buffers);
      if (mode === 'append') {
        onApplyBytes(merged, 'Merged PDFs');
      } else {
        const base = fileName.replace(/\.pdf$/i, '');
        await onCreateNewPdf(`${base} (Merged).pdf`, merged);
      }
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Merge failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleSplit = async () => {
    setBusy(true);
    try {
      const base = fileName.replace(/\.pdf$/i, '');
      if (splitMode === 'chunks') {
        const groups = splitIntoChunkSpecs(numPages, pagesPerChunk);
        for (let i = 0; i < groups.length; i++) {
          const out = await extractPdfPagesBytes(pdfBytes, groups[i]);
          await onCreateNewPdf(`${base} - Part ${i + 1}.pdf`, out);
        }
        onToast(`Created ${groups.length} split PDF files in vault`, 'success');
      } else {
        const indices = parsePageRange(rangeSpec, numPages);
        if (indices.length === 0) {
          onToast('Enter a valid page range (e.g. 1-3, 5).', 'warning');
          setBusy(false);
          return;
        }
        const out = await extractPdfPagesBytes(pdfBytes, indices);
        await onCreateNewPdf(`${base} (${rangeSpec}).pdf`, out);
      }
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Split failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleCompress = async () => {
    setBusy(true);
    try {
      const before = pdfBytes.byteLength;
      const compressed = await compressPdfWithQpdf(pdfBytes);
      const after = compressed.byteLength;
      onApplyBytes(
        compressed,
        `Compressed PDF (${formatBytes(before)} -> ${formatBytes(after)})`,
      );
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Compression failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleFlatten = async () => {
    setBusy(true);
    try {
      const flattened = await flattenPdfBytes(pdfBytes);
      onApplyBytes(flattened, 'Flattened interactive form fields into static PDF');
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Flatten failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleSaveMetadata = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    try {
      const updated = await updatePdfMetadata(pdfBytes, meta);
      onApplyBytes(updated, 'Updated PDF metadata');
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Failed to save metadata.', 'error');
    } finally {
      setBusy(false);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Organize PDF pages and properties"
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-3 sm:p-6"
    >
      <div className="flex max-h-[90vh] w-full max-w-5xl flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-2xl dark:border-slate-800 dark:bg-slate-900">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5 dark:border-slate-800">
          <div>
            <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
              Organize & Optimize PDF
            </h2>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              {fileName} — {numPages} pages ({formatBytes(pdfBytes.byteLength)})
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close Organize modal"
            className="flex h-10 w-10 items-center justify-center rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        {/* Section Tabs */}
        <div className="flex flex-wrap gap-1.5 border-b border-slate-200 bg-slate-50 px-5 py-2.5 dark:border-slate-800 dark:bg-slate-950">
          {(
            [
              ['pages', 'Pages & Reorder', Layers],
              ['merge', 'Merge PDFs', FilePlus],
              ['split', 'Split & Extract', Scissors],
              ['optimize', 'Compress, Flatten & Metadata', Minimize2],
            ] as const
          ).map(([id, label, Icon]) => (
            <button
              key={id}
              type="button"
              onClick={() => setSection(id)}
              className={`inline-flex min-h-[38px] items-center gap-2 whitespace-nowrap rounded-xl px-3.5 text-xs font-semibold transition-colors ${
                section === id
                  ? 'bg-blue-700 text-white dark:bg-blue-600'
                  : 'border border-slate-200 bg-white text-slate-700 hover:bg-slate-100 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-200'
              }`}
            >
              <Icon className="h-3.5 w-3.5" />
              {label}
            </button>
          ))}
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-5">
          {section === 'pages' && (
            <div className="space-y-4">
              {/* Action bar */}
              <div className="flex flex-wrap items-center gap-2 rounded-xl border border-slate-200 bg-slate-50 p-3 dark:border-slate-800 dark:bg-slate-950">
                <button
                  type="button"
                  onClick={() => rotateSelected(-90)}
                  className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                >
                  <RotateCcw className="h-3.5 w-3.5" /> Rotate Left
                </button>
                <button
                  type="button"
                  onClick={() => rotateSelected(90)}
                  className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                >
                  <RotateCw className="h-3.5 w-3.5" /> Rotate Right
                </button>
                <button
                  type="button"
                  onClick={() => insertBlank('blank-a4')}
                  className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                >
                  <FilePlus className="h-3.5 w-3.5" /> + Blank A4
                </button>
                <button
                  type="button"
                  onClick={() => insertBlank('blank-letter')}
                  className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                >
                  <FilePlus className="h-3.5 w-3.5" /> + Blank Letter
                </button>
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => void handleExtractSelected()}
                  className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200"
                >
                  <Scissors className="h-3.5 w-3.5" /> Extract Selected
                </button>
                <button
                  type="button"
                  onClick={deleteSelected}
                  className="inline-flex min-h-[38px] items-center gap-1.5 rounded-lg border border-rose-200 bg-rose-50 px-3 text-xs font-semibold text-rose-700 hover:bg-rose-100 dark:border-rose-900 dark:bg-rose-950/40 dark:text-rose-300"
                >
                  <Trash2 className="h-3.5 w-3.5" /> Delete Selected
                </button>
              </div>

              <p className="text-xs text-slate-500 dark:text-slate-400">
                Drag and drop cards or use arrow buttons to reorder pages. Click a card to select for rotation, extraction, or deletion.
              </p>

              <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5">
                {pages.map((p, idx) => (
                  <div
                    key={p.key}
                    draggable
                    onDragStart={() => setDragIdx(idx)}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={() => {
                      if (dragIdx !== null) movePage(dragIdx, idx);
                      setDragIdx(null);
                    }}
                    onClick={() => toggleSelect(idx)}
                    className={`group flex cursor-pointer flex-col items-center rounded-xl border p-3 transition-colors ${
                      p.selected
                        ? 'border-blue-600 bg-blue-50/70 dark:border-blue-500 dark:bg-blue-950/30'
                        : 'border-slate-200 bg-slate-50/70 hover:bg-slate-100 dark:border-slate-800 dark:bg-slate-950/60 dark:hover:bg-slate-800/70'
                    }`}
                  >
                    <div className="mb-2 flex w-full items-center justify-between text-[11px] font-semibold text-slate-600 dark:text-slate-300">
                      <span className="tabular-nums">#{idx + 1}</span>
                      <span
                        className={`flex h-5 w-5 items-center justify-center rounded border ${
                          p.selected
                            ? 'border-blue-600 bg-blue-600 text-white'
                            : 'border-slate-300 bg-white dark:border-slate-700 dark:bg-slate-900'
                        }`}
                      >
                        {p.selected && <Check className="h-3 w-3" />}
                      </span>
                    </div>

                    <MiniPageThumb
                      pdfDoc={pdfDoc}
                      pageNumber={typeof p.sourceIndex === 'number' ? p.sourceIndex + 1 : null}
                      rotationDeg={p.extraRotationDeg}
                    />

                    <span className="mt-2 truncate text-[11px] font-medium text-slate-700 dark:text-slate-300">
                      {p.label}
                      {p.extraRotationDeg !== 0 ? ` (+${p.extraRotationDeg}°)` : ''}
                    </span>

                    <div
                      className="mt-2 flex items-center gap-1"
                      onClick={(e) => e.stopPropagation()}
                    >
                      <button
                        type="button"
                        disabled={idx === 0}
                        onClick={() => movePage(idx, idx - 1)}
                        aria-label={`Move page ${idx + 1} up`}
                        className="flex h-7 w-7 items-center justify-center rounded border border-slate-200 bg-white text-slate-600 disabled:opacity-40 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300"
                      >
                        <ArrowUp className="h-3.5 w-3.5" />
                      </button>
                      <button
                        type="button"
                        disabled={idx === pages.length - 1}
                        onClick={() => movePage(idx, idx + 1)}
                        aria-label={`Move page ${idx + 1} down`}
                        className="flex h-7 w-7 items-center justify-center rounded border border-slate-200 bg-white text-slate-600 disabled:opacity-40 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300"
                      >
                        <ArrowDown className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {section === 'merge' && (
            <div className="space-y-5">
              <div>
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                  1. Select PDFs from your Local Vault
                </h3>
                <div className="mt-2 grid gap-2 sm:grid-cols-2">
                  {vaultPdfFiles.map((f) => {
                    const checked = mergeSelectedIds.includes(f.id);
                    return (
                      <label
                        key={f.id}
                        className={`flex cursor-pointer items-center justify-between rounded-xl border p-3 text-xs ${
                          checked
                            ? 'border-blue-600 bg-blue-50/70 dark:border-blue-500 dark:bg-blue-950/30'
                            : 'border-slate-200 dark:border-slate-800'
                        }`}
                      >
                        <span className="truncate font-semibold text-slate-800 dark:text-slate-200">
                          {f.name}
                        </span>
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() =>
                            setMergeSelectedIds((prev) =>
                              checked ? prev.filter((id) => id !== f.id) : [...prev, f.id],
                            )
                          }
                          className="h-4 w-4 accent-blue-600"
                        />
                      </label>
                    );
                  })}
                  {vaultPdfFiles.length === 0 && (
                    <p className="text-xs text-slate-500">
                      No other PDF files in local vault yet. Pick local files below.
                    </p>
                  )}
                </div>
              </div>

              <div>
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                  2. Or add PDF files from your device
                </h3>
                <input
                  type="file"
                  accept=".pdf,application/pdf"
                  multiple
                  onChange={(e) => setMergeLocalFiles(Array.from(e.target.files ?? []))}
                  className="mt-2 block w-full text-xs text-slate-600 file:mr-3 file:rounded-xl file:border-0 file:bg-blue-50 file:px-4 file:py-2 file:text-xs file:font-semibold file:text-blue-700 hover:file:bg-blue-100 dark:text-slate-300 dark:file:bg-blue-950 dark:file:text-blue-300"
                />
              </div>

              <div className="flex flex-wrap gap-3 pt-2">
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => void handleMerge('append')}
                  className="min-h-[42px] rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50 dark:bg-blue-600"
                >
                  Append into Current PDF
                </button>
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => void handleMerge('new')}
                  className="min-h-[42px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-800 hover:bg-slate-100 disabled:opacity-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                >
                  Save as New Merged PDF
                </button>
              </div>
            </div>
          )}

          {section === 'split' && (
            <div className="space-y-5">
              <div className="flex gap-3">
                <button
                  type="button"
                  onClick={() => setSplitMode('chunks')}
                  className={`min-h-[40px] rounded-xl px-4 text-xs font-semibold ${
                    splitMode === 'chunks'
                      ? 'bg-blue-700 text-white'
                      : 'border border-slate-200 text-slate-700 dark:border-slate-700 dark:text-slate-200'
                  }`}
                >
                  Fixed Pages per File
                </button>
                <button
                  type="button"
                  onClick={() => setSplitMode('range')}
                  className={`min-h-[40px] rounded-xl px-4 text-xs font-semibold ${
                    splitMode === 'range'
                      ? 'bg-blue-700 text-white'
                      : 'border border-slate-200 text-slate-700 dark:border-slate-700 dark:text-slate-200'
                  }`}
                >
                  Custom Page Range
                </button>
              </div>

              {splitMode === 'chunks' ? (
                <div className="max-w-xs space-y-2">
                  <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300">
                    Pages per split output file
                  </label>
                  <input
                    type="number"
                    min={1}
                    max={Math.max(1, numPages)}
                    value={pagesPerChunk}
                    onChange={(e) => setPagesPerChunk(Math.max(1, Number(e.target.value) || 1))}
                    className="h-10 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-sm tabular-nums dark:border-slate-700 dark:bg-slate-950"
                  />
                  <p className="text-xs text-slate-500">
                    Will generate {Math.ceil(numPages / Math.max(1, pagesPerChunk))} new PDF files in your local vault.
                  </p>
                </div>
              ) : (
                <div className="max-w-md space-y-2">
                  <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300">
                    Page range (e.g. 1-3, 5, 7-z)
                  </label>
                  <input
                    type="text"
                    value={rangeSpec}
                    onChange={(e) => setRangeSpec(e.target.value)}
                    placeholder="1-3, 5"
                    className="h-10 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-sm dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
              )}

              <button
                type="button"
                disabled={busy || numPages === 0}
                onClick={() => void handleSplit()}
                className="min-h-[42px] rounded-xl bg-blue-700 px-5 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50"
              >
                {busy ? 'Splitting...' : 'Split & Save to Local Vault'}
              </button>
            </div>
          )}

          {section === 'optimize' && (
            <div className="grid gap-6 lg:grid-cols-2">
              <div className="space-y-4 rounded-2xl border border-slate-200 p-4 dark:border-slate-800">
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                  Compress & Flatten
                </h3>
                <p className="text-xs text-slate-600 dark:text-slate-400">
                  Recompress object streams locally via qpdf WebAssembly or flatten interactive AcroForm fields into static page artwork.
                </p>
                <div className="flex flex-wrap gap-2.5">
                  <button
                    type="button"
                    disabled={busy}
                    onClick={() => void handleCompress()}
                    className="min-h-[40px] rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50"
                  >
                    {busy ? 'Working...' : 'Compress PDF Streams'}
                  </button>
                  <button
                    type="button"
                    disabled={busy}
                    onClick={() => void handleFlatten()}
                    className="min-h-[40px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                  >
                    Flatten Forms
                  </button>
                </div>
              </div>

              <form
                onSubmit={(e) => void handleSaveMetadata(e)}
                className="space-y-3 rounded-2xl border border-slate-200 p-4 dark:border-slate-800"
              >
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                  Document Metadata
                </h3>
                {(
                  [
                    ['title', 'Title'],
                    ['author', 'Author'],
                    ['subject', 'Subject'],
                    ['keywords', 'Keywords (comma-separated)'],
                  ] as const
                ).map(([field, label]) => (
                  <div key={field}>
                    <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-400">
                      {label}
                    </label>
                    <input
                      type="text"
                      value={meta[field]}
                      onChange={(e) => setMeta({...meta, [field]: e.target.value})}
                      className="mt-1 h-9 w-full rounded-lg border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
                    />
                  </div>
                ))}
                <button
                  type="submit"
                  disabled={busy}
                  className="min-h-[40px] rounded-xl bg-slate-900 px-4 text-xs font-semibold text-white hover:bg-slate-800 dark:bg-slate-100 dark:text-slate-900"
                >
                  Apply Metadata
                </button>
              </form>
            </div>
          )}
        </div>

        {/* Footer for Pages tab */}
        {section === 'pages' && (
          <div className="flex items-center justify-end gap-2.5 border-t border-slate-200 px-5 py-3 dark:border-slate-800">
            <button
              type="button"
              onClick={onClose}
              className="min-h-[40px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-700 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200"
            >
              Cancel
            </button>
            <button
              type="button"
              disabled={busy}
              onClick={() => void handleApplyPageOrder()}
              className="min-h-[40px] rounded-xl bg-blue-700 px-5 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50"
            >
              {busy ? 'Applying...' : 'Apply Page Changes'}
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
