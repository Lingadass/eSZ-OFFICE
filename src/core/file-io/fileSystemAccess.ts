import {MODULE_METADATA} from '../storage/fileRules';
import type {OfficeModuleType} from '../types';

interface OpenFilePickerOptionsLike {
  multiple?: boolean;
  types?: Array<{
    description?: string;
    accept: Record<string, string[]>;
  }>;
}

interface SaveFilePickerOptionsLike {
  suggestedName?: string;
  types?: Array<{
    description?: string;
    accept: Record<string, string[]>;
  }>;
}

interface FileSystemFileHandleLike {
  getFile: () => Promise<File>;
  createWritable: () => Promise<{
    write: (data: Blob | BufferSource | string) => Promise<void>;
    close: () => Promise<void>;
  }>;
}

export function isFileSystemAccessSupported(): boolean {
  return (
    typeof window !== 'undefined' &&
    'showOpenFilePicker' in window &&
    'showSaveFilePicker' in window
  );
}

export const SUPPORTED_OFFICE_ACCEPT =
  '.pdf,.docx,.doc,.pptx,.ppt,.xlsx,.xls,.csv,.txt,.md,.html,.rtf';

export async function pickLocalFiles(): Promise<File[]> {
  if (typeof window !== 'undefined' && 'showOpenFilePicker' in window) {
    try {
      const picker = (
        window as unknown as {
          showOpenFilePicker: (opts?: OpenFilePickerOptionsLike) => Promise<FileSystemFileHandleLike[]>;
        }
      ).showOpenFilePicker;
      const handles = await picker({
        multiple: true,
        types: [
          {
            description: 'Office & PDF Documents',
            accept: {
              'application/pdf': ['.pdf'],
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
              'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
              'text/csv': ['.csv'],
              'text/plain': ['.txt', '.md'],
            },
          },
        ],
      });
      return await Promise.all(handles.map((h) => h.getFile()));
    } catch (err) {
      if ((err as {name?: string})?.name === 'AbortError') {
        return [];
      }
      // Fall through to <input type="file"> on WebView / iframe restrictions
    }
  }

  return await new Promise<File[]>((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = SUPPORTED_OFFICE_ACCEPT;
    input.style.display = 'none';
    const cleanup = () => {
      if (input.parentNode) input.parentNode.removeChild(input);
    };
    input.addEventListener('change', () => {
      const list = input.files ? Array.from(input.files) : [];
      cleanup();
      resolve(list);
    });
    input.addEventListener('cancel', () => {
      cleanup();
      resolve([]);
    });
    document.body.appendChild(input);
    input.click();
  });
}

export async function saveBlobToDevice(
  blob: Blob,
  suggestedName: string,
  moduleType: OfficeModuleType,
): Promise<'fs-access' | 'download' | 'cancelled'> {
  const meta = MODULE_METADATA[moduleType];
  if (typeof window !== 'undefined' && 'showSaveFilePicker' in window) {
    try {
      const savePicker = (
        window as unknown as {
          showSaveFilePicker: (opts?: SaveFilePickerOptionsLike) => Promise<FileSystemFileHandleLike>;
        }
      ).showSaveFilePicker;
      const handle = await savePicker({
        suggestedName,
        types: [
          {
            description: meta.label,
            accept: {[meta.mimeType]: [meta.defaultExtension]},
          },
        ],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return 'fs-access';
    } catch (err) {
      if ((err as {name?: string})?.name === 'AbortError') {
        return 'cancelled';
      }
      // Fall through to anchor download when showSaveFilePicker is blocked in cross-origin iframe/WebView
    }
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = suggestedName;
  link.rel = 'noopener';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
  return 'download';
}

export function registerLaunchQueueConsumer(onFiles: (files: File[]) => void): () => void {
  if (typeof window === 'undefined' || !('launchQueue' in window)) {
    return () => undefined;
  }
  try {
    const lq = (
      window as unknown as {
        launchQueue: {
          setConsumer: (
            cb: (launchParams: {files?: FileSystemFileHandleLike[]}) => void,
          ) => void;
        };
      }
    ).launchQueue;
    lq.setConsumer(async (launchParams) => {
      if (!launchParams?.files?.length) return;
      const files = await Promise.all(launchParams.files.map((handle) => handle.getFile()));
      onFiles(files);
    });
  } catch {
    // Ignore unsupported launchQueue environments
  }
  return () => undefined;
}
