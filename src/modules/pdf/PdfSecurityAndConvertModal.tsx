import React, {useState} from 'react';
import {
  FileImage,
  FileText,
  Lock,
  ScanText,
  ShieldAlert,
  Unlock,
  X,
} from 'lucide-react';
import {saveBlobToDevice} from '../../core/file-io/fileSystemAccess';
import {
  applyTrueRedactions,
  decryptPdfWithQpdf,
  encryptPdfWithQpdf,
} from './pdfEngine';
import {
  imagesToPdfBytes,
  officeFileToPdfBestEffort,
  pdfToImagesZip,
} from './pdfConvert';
import {runOcrToSearchablePdf, type OcrProgressUpdate} from './pdfOcrWorker';
import {parsePageRange, textToPdfBytes} from './pdfPure';
import type {PdfAnnotationItem} from './pdfTypes';

interface PdfSecurityAndConvertModalProps {
  open: boolean;
  initialTab?: 'security' | 'convert' | 'ocr';
  onClose: () => void;
  fileName: string;
  numPages: number;
  pdfBytes: Uint8Array;
  currentPassword?: string;
  pendingRedactions: PdfAnnotationItem[];
  onApplyBytes: (nextBytes: Uint8Array, summary: string, nextPassword?: string) => void;
  onClearRedactions: () => void;
  onCreateNewPdf: (name: string, bytes: Uint8Array) => Promise<void>;
  onToast: (msg: string, tone?: 'info' | 'success' | 'warning' | 'error') => void;
}

export const PdfSecurityAndConvertModal: React.FC<PdfSecurityAndConvertModalProps> = ({
  open,
  initialTab = 'security',
  onClose,
  fileName,
  numPages,
  pdfBytes,
  currentPassword,
  pendingRedactions,
  onApplyBytes,
  onClearRedactions,
  onCreateNewPdf,
  onToast,
}) => {
  const [tab, setTab] = useState<'security' | 'convert' | 'ocr'>(initialTab);
  const [busy, setBusy] = useState(false);

  // Security states
  const [userPassword, setUserPassword] = useState('');
  const [ownerPassword, setOwnerPassword] = useState('');
  const [allowPrint, setAllowPrint] = useState(true);
  const [allowModify, setAllowModify] = useState(false);
  const [allowExtract, setAllowExtract] = useState(false);
  const [allowAnnotate, setAllowAnnotate] = useState(true);
  const [decryptPassword, setDecryptPassword] = useState(currentPassword ?? '');

  // Convert states
  const [exportImgFormat, setExportImgFormat] = useState<'png' | 'jpeg'>('png');
  const [exportImgProgress, setExportImgProgress] = useState<string>('');
  const [plainTextTitle, setPlainTextTitle] = useState('');
  const [plainTextBody, setPlainTextBody] = useState('');

  // OCR states
  const [ocrRange, setOcrRange] = useState('');
  const [ocrProgress, setOcrProgress] = useState<OcrProgressUpdate | null>(null);

  React.useEffect(() => {
    if (open) {
      setTab(initialTab);
      setDecryptPassword(currentPassword ?? '');
      setOcrProgress(null);
      setExportImgProgress('');
    }
  }, [open, initialTab, currentPassword]);

  if (!open) return null;

  const baseName = fileName.replace(/\.pdf$/i, '');

  const handleEncrypt = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!userPassword.trim()) {
      onToast('Enter an open password to encrypt this PDF.', 'warning');
      return;
    }
    setBusy(true);
    try {
      const encrypted = await encryptPdfWithQpdf(pdfBytes, {
        userPassword: userPassword.trim(),
        ownerPassword: ownerPassword.trim() || userPassword.trim(),
        allowPrint,
        allowModify,
        allowExtract,
        allowAnnotate,
      });
      onApplyBytes(encrypted, 'Encrypted PDF with AES-256 (qpdf)', userPassword.trim());
      setUserPassword('');
      setOwnerPassword('');
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'PDF encryption failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleDecrypt = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!decryptPassword) {
      onToast('Enter the current PDF password to remove protection.', 'warning');
      return;
    }
    setBusy(true);
    try {
      const decrypted = await decryptPdfWithQpdf(pdfBytes, decryptPassword);
      onApplyBytes(decrypted, 'Removed PDF password protection', undefined);
      onClose();
    } catch (err) {
      onToast(
        err instanceof Error
          ? `Decryption failed: ${err.message}`
          : 'Incorrect password or decryption failed.',
        'error',
      );
    } finally {
      setBusy(false);
    }
  };

  const handleCommitTrueRedactions = async () => {
    if (pendingRedactions.length === 0) {
      onToast('Draw at least one Redact box on the PDF first.', 'info');
      return;
    }
    setBusy(true);
    try {
      const redactedBytes = await applyTrueRedactions(
        pdfBytes,
        pendingRedactions,
        currentPassword,
      );
      onClearRedactions();
      onApplyBytes(
        redactedBytes,
        `Permanently redacted ${pendingRedactions.length} region(s)`,
        currentPassword,
      );
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Redaction failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleImagesToPdf = async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    setBusy(true);
    try {
      const bytes = await imagesToPdfBytes(Array.from(files));
      const first = files[0].name.replace(/\.[^.]+$/, '');
      await onCreateNewPdf(`${first} (Images).pdf`, bytes);
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Image to PDF conversion failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handlePdfToImages = async () => {
    setBusy(true);
    setExportImgProgress('Starting...');
    try {
      const {zipBlob, pageCount} = await pdfToImagesZip(
        pdfBytes,
        exportImgFormat,
        2,
        (page, total) => setExportImgProgress(`Rendering page ${page} of ${total}...`),
      );
      const outName = `${baseName}-pages-${exportImgFormat}.zip`;
      await saveBlobToDevice(zipBlob, outName, 'pdf');
      onToast(`Exported ${pageCount} ${exportImgFormat.toUpperCase()} images in ${outName}`, 'success');
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'PDF to images export failed.', 'error');
    } finally {
      setBusy(false);
      setExportImgProgress('');
    }
  };

  const handleOfficeToPdf = async (file: File | undefined) => {
    if (!file) return;
    setBusy(true);
    try {
      const bytes = await officeFileToPdfBestEffort(file);
      const cleanName = file.name.replace(/\.[^.]+$/, '');
      await onCreateNewPdf(`${cleanName}.pdf`, bytes);
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Conversion failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handlePlainTextToPdf = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!plainTextBody.trim()) {
      onToast('Enter some text to convert to PDF.', 'info');
      return;
    }
    setBusy(true);
    try {
      const title = plainTextTitle.trim() || 'Text Document';
      const bytes = await textToPdfBytes(plainTextBody, title);
      await onCreateNewPdf(`${title}.pdf`, bytes);
      setPlainTextTitle('');
      setPlainTextBody('');
      onClose();
    } catch (err) {
      onToast(err instanceof Error ? err.message : 'Text to PDF failed.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleRunOcr = async () => {
    setBusy(true);
    try {
      const indices = ocrRange.trim() ? parsePageRange(ocrRange, numPages) : [];
      const {pdfBytes: searchableBytes} = await runOcrToSearchablePdf(
        pdfBytes,
        indices,
        'eng',
        (info) => setOcrProgress(info),
      );
      await onCreateNewPdf(`${baseName} (Searchable OCR).pdf`, searchableBytes);
      onClose();
    } catch (err) {
      onToast(
        err instanceof Error
          ? `OCR failed: ${err.message}`
          : 'OCR requires cached or reachable English traineddata.',
        'error',
      );
    } finally {
      setBusy(false);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="PDF Security, Conversion, and OCR"
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 p-3 sm:p-6"
    >
      <div className="flex max-h-[90vh] w-full max-w-4xl flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-2xl dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5 dark:border-slate-800">
          <div>
            <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
              Security, Convert & OCR
            </h2>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              All cryptographic encryption, redaction, conversion, and OCR run locally on your device.
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close Security, Convert and OCR modal"
            className="flex h-10 w-10 items-center justify-center rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-100 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="flex flex-wrap gap-1.5 border-b border-slate-200 bg-slate-50 px-5 py-2.5 dark:border-slate-800 dark:bg-slate-950">
          {(
            [
              ['security', 'Security & Redact', Lock],
              ['convert', 'Convert Files', FileImage],
              ['ocr', 'OCR to Searchable PDF', ScanText],
            ] as const
          ).map(([id, label, Icon]) => (
            <button
              key={id}
              type="button"
              onClick={() => setTab(id)}
              className={`inline-flex min-h-[38px] items-center gap-2 whitespace-nowrap rounded-xl px-4 text-xs font-semibold ${
                tab === id
                  ? 'bg-blue-700 text-white'
                  : 'border border-slate-200 bg-white text-slate-700 hover:bg-slate-100 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-200'
              }`}
            >
              <Icon className="h-3.5 w-3.5" />
              {label}
            </button>
          ))}
        </div>

        <div className="flex-1 overflow-y-auto p-5">
          {tab === 'security' && (
            <div className="grid gap-5 lg:grid-cols-2">
              {/* Encrypt */}
              <form
                onSubmit={(e) => void handleEncrypt(e)}
                className="space-y-3 rounded-2xl border border-slate-200 p-4 dark:border-slate-800"
              >
                <div className="flex items-center gap-2">
                  <Lock className="h-4 w-4 text-blue-700 dark:text-blue-400" />
                  <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                    Set Password & Permissions (AES-256)
                  </h3>
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-400">
                    Document Open Password
                  </label>
                  <input
                    type="password"
                    value={userPassword}
                    onChange={(e) => setUserPassword(e.target.value)}
                    placeholder="Required to open PDF"
                    className="mt-1 h-10 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-400">
                    Owner / Permissions Password (optional)
                  </label>
                  <input
                    type="password"
                    value={ownerPassword}
                    onChange={(e) => setOwnerPassword(e.target.value)}
                    placeholder="Defaults to open password"
                    className="mt-1 h-10 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
                  />
                </div>
                <div className="grid grid-cols-2 gap-2 pt-1 text-xs text-slate-700 dark:text-slate-300">
                  <label className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={allowPrint}
                      onChange={(e) => setAllowPrint(e.target.checked)}
                      className="h-4 w-4 accent-blue-600"
                    />
                    Allow Printing
                  </label>
                  <label className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={allowModify}
                      onChange={(e) => setAllowModify(e.target.checked)}
                      className="h-4 w-4 accent-blue-600"
                    />
                    Allow Editing
                  </label>
                  <label className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={allowExtract}
                      onChange={(e) => setAllowExtract(e.target.checked)}
                      className="h-4 w-4 accent-blue-600"
                    />
                    Allow Copying
                  </label>
                  <label className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={allowAnnotate}
                      onChange={(e) => setAllowAnnotate(e.target.checked)}
                      className="h-4 w-4 accent-blue-600"
                    />
                    Allow Annotations
                  </label>
                </div>
                <button
                  type="submit"
                  disabled={busy}
                  className="min-h-[42px] w-full rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50"
                >
                  {busy ? 'Encrypting...' : 'Encrypt PDF'}
                </button>
              </form>

              <div className="space-y-5">
                {/* Decrypt */}
                <form
                  onSubmit={(e) => void handleDecrypt(e)}
                  className="space-y-3 rounded-2xl border border-slate-200 p-4 dark:border-slate-800"
                >
                  <div className="flex items-center gap-2">
                    <Unlock className="h-4 w-4 text-emerald-700 dark:text-emerald-400" />
                    <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                      Remove Password
                    </h3>
                  </div>
                  <input
                    type="password"
                    value={decryptPassword}
                    onChange={(e) => setDecryptPassword(e.target.value)}
                    placeholder="Enter current password..."
                    className="h-10 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
                  />
                  <button
                    type="submit"
                    disabled={busy}
                    className="min-h-[40px] rounded-xl border border-slate-200 px-4 text-xs font-semibold text-slate-800 hover:bg-slate-100 disabled:opacity-50 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-800"
                  >
                    Decrypt & Remove Password
                  </button>
                </form>

                {/* True Redaction */}
                <div className="space-y-3 rounded-2xl border border-rose-200 bg-rose-50/60 p-4 dark:border-rose-900/60 dark:bg-rose-950/20">
                  <div className="flex items-center gap-2">
                    <ShieldAlert className="h-4 w-4 text-rose-700 dark:text-rose-400" />
                    <h3 className="text-sm font-bold text-rose-950 dark:text-rose-200">
                      Permanent Redaction ({pendingRedactions.length} marked)
                    </h3>
                  </div>
                  <p className="text-xs leading-relaxed text-rose-900/80 dark:text-rose-300/90">
                    Use the <strong>Redact</strong> tool on the canvas to mark sensitive areas, then commit here. Affected pages are rasterized at 2x resolution with solid black fills so underlying selectable text is irreversibly destroyed.
                  </p>
                  <button
                    type="button"
                    disabled={busy || pendingRedactions.length === 0}
                    onClick={() => void handleCommitTrueRedactions()}
                    className="min-h-[40px] rounded-xl bg-rose-700 px-4 text-xs font-semibold text-white hover:bg-rose-800 disabled:opacity-50"
                  >
                    {busy ? 'Redacting...' : 'Apply Permanent Redactions'}
                  </button>
                </div>
              </div>
            </div>
          )}

          {tab === 'convert' && (
            <div className="grid gap-5 lg:grid-cols-2">
              {/* PDF -> Images */}
              <div className="space-y-3 rounded-2xl border border-slate-200 p-4 dark:border-slate-800">
                <div className="flex items-center gap-2">
                  <FileImage className="h-4 w-4 text-blue-700 dark:text-blue-400" />
                  <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                    Export Current PDF as Images (.zip)
                  </h3>
                </div>
                <div className="flex gap-2">
                  {(['png', 'jpeg'] as const).map((fmt) => (
                    <button
                      key={fmt}
                      type="button"
                      onClick={() => setExportImgFormat(fmt)}
                      className={`min-h-[36px] rounded-lg px-3.5 text-xs font-semibold uppercase ${
                        exportImgFormat === fmt
                          ? 'bg-slate-900 text-white dark:bg-slate-100 dark:text-slate-900'
                          : 'border border-slate-200 text-slate-700 dark:border-slate-700 dark:text-slate-300'
                      }`}
                    >
                      {fmt}
                    </button>
                  ))}
                </div>
                {exportImgProgress && (
                  <p className="text-xs font-medium text-blue-700 dark:text-blue-400">
                    {exportImgProgress}
                  </p>
                )}
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => void handlePdfToImages()}
                  className="min-h-[40px] rounded-xl bg-blue-700 px-4 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50"
                >
                  Export {numPages} Page{numPages === 1 ? '' : 's'} to ZIP
                </button>
              </div>

              {/* Images -> PDF */}
              <div className="space-y-3 rounded-2xl border border-slate-200 p-4 dark:border-slate-800">
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                  Images to PDF (PNG, JPG, WebP)
                </h3>
                <p className="text-xs text-slate-500 dark:text-slate-400">
                  Combines selected images into a new multi-page PDF saved in your local vault.
                </p>
                <input
                  type="file"
                  accept="image/png,image/jpeg,image/webp"
                  multiple
                  disabled={busy}
                  onChange={(e) => void handleImagesToPdf(e.target.files)}
                  className="block w-full text-xs text-slate-600 file:mr-3 file:rounded-xl file:border-0 file:bg-blue-50 file:px-4 file:py-2 file:text-xs file:font-semibold file:text-blue-700 dark:text-slate-300 dark:file:bg-blue-950 dark:file:text-blue-300"
                />
              </div>

              {/* Word/PPT -> PDF (Best Effort) */}
              <div className="space-y-3 rounded-2xl border border-slate-200 p-4 dark:border-slate-800">
                <div className="flex items-center gap-2">
                  <FileText className="h-4 w-4 text-blue-700 dark:text-blue-400" />
                  <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                    Word / PowerPoint / Text File to PDF
                  </h3>
                </div>
                <p className="text-xs text-slate-500 dark:text-slate-400">
                  Best-effort on-device conversion for <code>.docx</code>, <code>.pptx</code>, <code>.txt</code>, <code>.md</code>, and <code>.csv</code>. Extracts text structure cleanly; complex Word/Slide floating layouts are simplified.
                </p>
                <input
                  type="file"
                  accept=".docx,.pptx,.txt,.md,.csv"
                  disabled={busy}
                  onChange={(e) => void handleOfficeToPdf(e.target.files?.[0])}
                  className="block w-full text-xs text-slate-600 file:mr-3 file:rounded-xl file:border-0 file:bg-slate-100 file:px-4 file:py-2 file:text-xs file:font-semibold file:text-slate-800 dark:text-slate-300 dark:file:bg-slate-800 dark:file:text-slate-200"
                />
              </div>

              {/* Direct Text -> PDF */}
              <form
                onSubmit={(e) => void handlePlainTextToPdf(e)}
                className="space-y-2.5 rounded-2xl border border-slate-200 p-4 dark:border-slate-800"
              >
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                  Quick Text to PDF
                </h3>
                <input
                  type="text"
                  value={plainTextTitle}
                  onChange={(e) => setPlainTextTitle(e.target.value)}
                  placeholder="Document title..."
                  className="h-9 w-full rounded-lg border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
                <textarea
                  rows={3}
                  value={plainTextBody}
                  onChange={(e) => setPlainTextBody(e.target.value)}
                  placeholder="Paste or type text content..."
                  className="w-full rounded-lg border border-slate-200 bg-slate-50 p-2.5 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
                <button
                  type="submit"
                  disabled={busy}
                  className="min-h-[38px] rounded-xl bg-slate-900 px-4 text-xs font-semibold text-white hover:bg-slate-800 dark:bg-slate-100 dark:text-slate-900"
                >
                  Create PDF from Text
                </button>
              </form>
            </div>
          )}

          {tab === 'ocr' && (
            <div className="space-y-4 rounded-2xl border border-slate-200 p-5 dark:border-slate-800">
              <div className="flex items-start gap-3">
                <ScanText className="mt-0.5 h-5 w-5 shrink-0 text-blue-700 dark:text-blue-400" />
                <div>
                  <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                    Scanned PDF to Searchable PDF (Lazy Tesseract.js Worker)
                  </h3>
                  <p className="mt-1 text-xs leading-relaxed text-slate-600 dark:text-slate-400">
                    Recognizes English text on scanned PDF pages completely inside a local browser worker and embeds an invisible, selectable text layer into a new searchable PDF in your vault.
                  </p>
                  <p className="mt-1.5 text-xs text-amber-700 dark:text-amber-400">
                    Note: The first OCR run downloads the English traineddata model (~11 MB) and caches it in the Service Worker for future offline use.
                  </p>
                </div>
              </div>

              <div className="max-w-sm space-y-1.5">
                <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300">
                  Pages to OCR (leave blank for all {numPages} pages)
                </label>
                <input
                  type="text"
                  value={ocrRange}
                  onChange={(e) => setOcrRange(e.target.value)}
                  placeholder={`e.g. 1-${Math.min(3, Math.max(1, numPages))}`}
                  className="h-10 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs dark:border-slate-700 dark:bg-slate-950"
                />
              </div>

              {ocrProgress && (
                <div className="space-y-1.5 rounded-xl border border-blue-200 bg-blue-50/70 p-3 dark:border-blue-900 dark:bg-blue-950/30">
                  <div className="flex items-center justify-between text-xs font-semibold text-blue-900 dark:text-blue-200">
                    <span>
                      Page {ocrProgress.page} of {ocrProgress.totalPages} — {ocrProgress.status}
                    </span>
                    <span className="tabular-nums">{ocrProgress.percent}%</span>
                  </div>
                  <div className="h-2 overflow-hidden rounded-full bg-blue-200 dark:bg-blue-900">
                    <div
                      className="h-full bg-blue-700 transition-all dark:bg-blue-500"
                      style={{width: `${Math.min(100, Math.max(5, ocrProgress.percent))}%`}}
                    />
                  </div>
                </div>
              )}

              <button
                type="button"
                disabled={busy || numPages === 0}
                onClick={() => void handleRunOcr()}
                className="min-h-[42px] rounded-xl bg-blue-700 px-5 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50"
              >
                {busy ? 'Running On-Device OCR...' : 'Run OCR & Save Searchable PDF'}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
