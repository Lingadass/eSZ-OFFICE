import React, {useEffect, useRef, useState} from 'react';
import {CheckSquare, Eraser, PenTool, Type, Upload, X} from 'lucide-react';
import {applyPdfFormFields, getPdfFormFields} from './pdfEngine';
import type {PdfFormFieldInfo} from './pdfTypes';

interface PdfFormsAndSignModalProps {
  open: boolean;
  initialTab?: 'forms' | 'signature';
  onClose: () => void;
  pdfBytes: Uint8Array;
  onApplyBytes: (nextBytes: Uint8Array, summary: string) => void;
  onReadySignature: (dataUrl: string, placeImmediately: boolean) => void;
  onToast: (msg: string, tone?: 'info' | 'success' | 'warning' | 'error') => void;
}

export const PdfFormsAndSignModal: React.FC<PdfFormsAndSignModalProps> = ({
  open,
  initialTab = 'signature',
  onClose,
  pdfBytes,
  onApplyBytes,
  onReadySignature,
  onToast,
}) => {
  const [tab, setTab] = useState<'forms' | 'signature'>(initialTab);
  const [fields, setFields] = useState<PdfFormFieldInfo[]>([]);
  const [loadingFields, setLoadingFields] = useState(false);
  const [flattenAfter, setFlattenAfter] = useState(false);

  // Signature state
  const [signMode, setSignMode] = useState<'draw' | 'type' | 'upload'>('draw');
  const [inkColor, setInkColor] = useState('#0F172A');
  const [typedName, setTypedName] = useState('');
  const [uploadedUrl, setUploadedUrl] = useState<string | null>(null);
  const [hasDrawn, setHasDrawn] = useState(false);

  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const drawingRef = useRef(false);

  useEffect(() => {
    if (!open) return;
    setTab(initialTab);
    setLoadingFields(true);
    void getPdfFormFields(pdfBytes)
      .then(setFields)
      .catch(() => setFields([]))
      .finally(() => setLoadingFields(false));
  }, [open, initialTab, pdfBytes]);

  useEffect(() => {
    if (!open || tab !== 'signature' || signMode !== 'draw') return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
  }, [open, tab, signMode]);

  if (!open) return null;

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    ctx?.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
  };

  const getCanvasPoint = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current!;
    const rect = canvas.getBoundingClientRect();
    return {
      x: ((e.clientX - rect.left) / rect.width) * canvas.width,
      y: ((e.clientY - rect.top) / rect.height) * canvas.height,
    };
  };

  const handleDrawDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    e.currentTarget.setPointerCapture(e.pointerId);
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext('2d');
    if (!ctx) return;
    const pt = getCanvasPoint(e);
    drawingRef.current = true;
    ctx.strokeStyle = inkColor;
    ctx.lineWidth = 3;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.beginPath();
    ctx.moveTo(pt.x, pt.y);
    setHasDrawn(true);
  };

  const handleDrawMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!drawingRef.current) return;
    const ctx = canvasRef.current?.getContext('2d');
    if (!ctx) return;
    const pt = getCanvasPoint(e);
    ctx.lineTo(pt.x, pt.y);
    ctx.stroke();
  };

  const handleDrawUp = () => {
    drawingRef.current = false;
  };

  const renderTypedSignaturePng = (name: string): string | null => {
    if (!name.trim()) return null;
    const off = document.createElement('canvas');
    off.width = 560;
    off.height = 180;
    const ctx = off.getContext('2d');
    if (!ctx) return null;
    ctx.clearRect(0, 0, off.width, off.height);
    ctx.fillStyle = inkColor;
    ctx.font = 'italic 58px Georgia, "Times New Roman", serif';
    ctx.textBaseline = 'middle';
    ctx.fillText(name.trim(), 24, 92, 512);
    return off.toDataURL('image/png');
  };

  const handleSignatureUpload = (file: File | undefined) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        setUploadedUrl(reader.result);
      }
    };
    reader.readAsDataURL(file);
  };

  const getActiveSignatureUrl = (): string | null => {
    if (signMode === 'draw') {
      if (!hasDrawn || !canvasRef.current) return null;
      return canvasRef.current.toDataURL('image/png');
    }
    if (signMode === 'type') {
      return renderTypedSignaturePng(typedName);
    }
    return uploadedUrl;
  };

  const handleConfirmSignature = (placeImmediately: boolean) => {
    const url = getActiveSignatureUrl();
    if (!url) {
      onToast('Create or upload a signature first.', 'info');
      return;
    }
    onReadySignature(url, placeImmediately);
    onClose();
  };

  const handleApplyForm = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      const next = await applyPdfFormFields(pdfBytes, fields, flattenAfter);
      onApplyBytes(next, flattenAfter ? 'Filled and flattened PDF form' : 'Updated PDF form fields');
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Failed to update form fields.', 'error');
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="PDF Forms and Signature"
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-3 sm:p-6"
    >
      <div className="flex max-h-[90vh] w-full max-w-2xl flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-2xl dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5 dark:border-slate-800">
          <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
            Forms & Signature
          </h2>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close Forms and Signature modal"
            className="flex h-10 w-10 items-center justify-center rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="flex gap-2 border-b border-slate-200 bg-slate-50 px-5 py-2.5 dark:border-slate-800 dark:bg-slate-950">
          <button
            type="button"
            onClick={() => setTab('signature')}
            className={`inline-flex min-h-[38px] items-center gap-2 rounded-xl px-4 text-xs font-semibold ${
              tab === 'signature'
                ? 'bg-blue-700 text-white'
                : 'border border-slate-200 bg-white text-slate-700 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-200'
            }`}
          >
            <PenTool className="h-3.5 w-3.5" /> Add Signature
          </button>
          <button
            type="button"
            onClick={() => setTab('forms')}
            className={`inline-flex min-h-[38px] items-center gap-2 rounded-xl px-4 text-xs font-semibold ${
              tab === 'forms'
                ? 'bg-blue-700 text-white'
                : 'border border-slate-200 bg-white text-slate-700 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-200'
            }`}
          >
            <CheckSquare className="h-3.5 w-3.5" /> Fill Form Fields ({fields.length})
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-5">
          {tab === 'signature' && (
            <div className="space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div className="flex gap-1.5">
                  {(
                    [
                      ['draw', 'Draw', PenTool],
                      ['type', 'Type', Type],
                      ['upload', 'Upload', Upload],
                    ] as const
                  ).map(([id, label, Icon]) => (
                    <button
                      key={id}
                      type="button"
                      onClick={() => setSignMode(id)}
                      className={`inline-flex min-h-[36px] items-center gap-1.5 rounded-lg px-3 text-xs font-semibold ${
                        signMode === id
                          ? 'bg-slate-900 text-white dark:bg-slate-100 dark:text-slate-900'
                          : 'border border-slate-200 text-slate-700 dark:border-slate-700 dark:text-slate-300'
                      }`}
                    >
                      <Icon className="h-3.5 w-3.5" />
                      {label}
                    </button>
                  ))}
                </div>

                {signMode !== 'upload' && (
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-medium text-slate-500">Ink:</span>
                    {['#0F172A', '#1D4ED8', '#DC2626'].map((c) => (
                      <button
                        key={c}
                        type="button"
                        onClick={() => setInkColor(c)}
                        aria-label={`Ink color ${c}`}
                        className={`h-6 w-6 rounded-full border-2 ${
                          inkColor === c ? 'border-blue-500 scale-110' : 'border-slate-300'
                        }`}
                        style={{backgroundColor: c}}
                      />
                    ))}
                  </div>
                )}
              </div>

              {signMode === 'draw' && (
                <div className="space-y-2">
                  <div className="relative overflow-hidden rounded-xl border border-slate-300 bg-white dark:border-slate-700">
                    <canvas
                      ref={canvasRef}
                      width={560}
                      height={180}
                      onPointerDown={handleDrawDown}
                      onPointerMove={handleDrawMove}
                      onPointerUp={handleDrawUp}
                      onPointerCancel={handleDrawUp}
                      className="h-44 w-full touch-none cursor-crosshair"
                    />
                    <button
                      type="button"
                      onClick={clearCanvas}
                      className="absolute right-2.5 top-2.5 inline-flex min-h-[32px] items-center gap-1 rounded-lg border border-slate-200 bg-white/90 px-2.5 text-xs font-medium text-slate-700 hover:bg-slate-100"
                    >
                      <Eraser className="h-3.5 w-3.5" /> Clear
                    </button>
                  </div>
                  <p className="text-xs text-slate-500">
                    Draw your signature with mouse, stylus, or finger above. Transparent PNG output.
                  </p>
                </div>
              )}

              {signMode === 'type' && (
                <div className="space-y-3">
                  <input
                    type="text"
                    value={typedName}
                    onChange={(e) => setTypedName(e.target.value)}
                    placeholder="Type your full name..."
                    className="h-11 w-full rounded-xl border border-slate-200 bg-slate-50 px-3.5 text-sm dark:border-slate-700 dark:bg-slate-950"
                  />
                  <div className="flex h-28 items-center justify-center rounded-xl border border-slate-200 bg-white px-4">
                    <span
                      className="font-serif text-4xl italic"
                      style={{color: inkColor}}
                    >
                      {typedName.trim() || 'Signature Preview'}
                    </span>
                  </div>
                </div>
              )}

              {signMode === 'upload' && (
                <div className="space-y-3">
                  <input
                    type="file"
                    accept="image/png,image/jpeg"
                    onChange={(e) => handleSignatureUpload(e.target.files?.[0])}
                    className="block w-full text-xs text-slate-600 file:mr-3 file:rounded-xl file:border-0 file:bg-blue-50 file:px-4 file:py-2 file:text-xs font-semibold file:text-blue-700 dark:text-slate-300"
                  />
                  {uploadedUrl && (
                    <div className="flex h-32 items-center justify-center rounded-xl border border-slate-200 bg-white p-3">
                      <img src={uploadedUrl} alt="Uploaded signature" className="max-h-full object-contain" />
                    </div>
                  )}
                </div>
              )}

              <div className="flex flex-wrap justify-end gap-2.5 pt-2">
                <button
                  type="button"
                  onClick={() => handleConfirmSignature(true)}
                  className="min-h-[42px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-800 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-200"
                >
                  Stamp on Current Page
                </button>
                <button
                  type="button"
                  onClick={() => handleConfirmSignature(false)}
                  className="min-h-[42px] rounded-xl bg-blue-700 px-5 text-xs font-semibold text-white hover:bg-blue-800"
                >
                  Click Page to Place Signature
                </button>
              </div>
            </div>
          )}

          {tab === 'forms' && (
            <div>
              {loadingFields ? (
                <p className="text-xs text-slate-500">Scanning PDF for interactive AcroForm fields...</p>
              ) : fields.length === 0 ? (
                <div className="rounded-xl border border-slate-200 bg-slate-50 p-5 text-center dark:border-slate-800 dark:bg-slate-950">
                  <p className="text-sm font-semibold text-slate-800 dark:text-slate-200">
                    No interactive AcroForm fields found in this PDF.
                  </p>
                  <p className="mt-1.5 text-xs text-slate-500 dark:text-slate-400">
                    For flat scanned or printed forms, use the Text Box tool and Add Signature tab to fill and sign directly on the page.
                  </p>
                </div>
              ) : (
                <form onSubmit={(e) => void handleApplyForm(e)} className="space-y-4">
                  <div className="space-y-3">
                    {fields.map((f, idx) => (
                      <div
                        key={f.name}
                        className="rounded-xl border border-slate-200 p-3 dark:border-slate-800"
                      >
                        <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300">
                          {f.name}
                        </label>
                        {f.type === 'checkbox' ? (
                          <label className="mt-2 inline-flex items-center gap-2 text-xs">
                            <input
                              type="checkbox"
                              checked={Boolean(f.value)}
                              onChange={(e) =>
                                setFields((prev) =>
                                  prev.map((item, i) =>
                                    i === idx ? {...item, value: e.target.checked} : item,
                                  ),
                                )
                              }
                              className="h-4 w-4 accent-blue-600"
                            />
                            Checked
                          </label>
                        ) : (f.type === 'dropdown' || f.type === 'radio') && f.options ? (
                          <select
                            value={String(f.value)}
                            onChange={(e) =>
                              setFields((prev) =>
                                prev.map((item, i) =>
                                  i === idx ? {...item, value: e.target.value} : item,
                                ),
                              )
                            }
                            className="mt-1.5 h-10 w-full rounded-lg border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
                          >
                            <option value="">Select option...</option>
                            {f.options.map((opt) => (
                              <option key={opt} value={opt}>
                                {opt}
                              </option>
                            ))}
                          </select>
                        ) : (
                          <input
                            type="text"
                            value={String(f.value)}
                            onChange={(e) =>
                              setFields((prev) =>
                                prev.map((item, i) =>
                                  i === idx ? {...item, value: e.target.value} : item,
                                ),
                              )
                            }
                            className="mt-1.5 h-10 w-full rounded-lg border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
                          />
                        )}
                      </div>
                    ))}
                  </div>

                  <label className="flex items-center gap-2 text-xs font-medium text-slate-700 dark:text-slate-300">
                    <input
                      type="checkbox"
                      checked={flattenAfter}
                      onChange={(e) => setFlattenAfter(e.target.checked)}
                      className="h-4 w-4 accent-blue-600"
                    />
                    Flatten form fields permanently after saving
                  </label>

                  <button
                    type="submit"
                    className="min-h-[42px] rounded-xl bg-blue-700 px-5 text-xs font-semibold text-white hover:bg-blue-800"
                  >
                    Apply Form Values
                  </button>
                </form>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
