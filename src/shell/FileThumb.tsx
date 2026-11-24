import React from 'react';
import {FileSpreadsheet, FileText, Presentation} from 'lucide-react';
import type {OfficeModuleType} from '../core/types';

export const ModuleIcon: React.FC<{type: OfficeModuleType; size?: 'sm' | 'md' | 'lg'}> = ({
  type,
  size = 'md',
}) => {
  const containerClass =
    size === 'sm'
      ? 'h-8 w-8 rounded-lg'
      : size === 'lg'
        ? 'h-12 w-12 rounded-xl'
        : 'h-10 w-10 rounded-xl';
  const iconClass = size === 'sm' ? 'h-4 w-4' : size === 'lg' ? 'h-6 w-6' : 'h-5 w-5';

  switch (type) {
    case 'pdf':
      return (
        <div
          className={`flex shrink-0 items-center justify-center bg-red-50 text-red-700 dark:bg-red-950/60 dark:text-red-300 ${containerClass}`}
        >
          <FileText className={iconClass} />
        </div>
      );
    case 'docs':
      return (
        <div
          className={`flex shrink-0 items-center justify-center bg-blue-50 text-blue-700 dark:bg-blue-950/60 dark:text-blue-300 ${containerClass}`}
        >
          <FileText className={iconClass} />
        </div>
      );
    case 'sheets':
      return (
        <div
          className={`flex shrink-0 items-center justify-center bg-emerald-50 text-emerald-700 dark:bg-emerald-950/60 dark:text-emerald-300 ${containerClass}`}
        >
          <FileSpreadsheet className={iconClass} />
        </div>
      );
    case 'slides':
      return (
        <div
          className={`flex shrink-0 items-center justify-center bg-amber-50 text-amber-700 dark:bg-amber-950/60 dark:text-amber-300 ${containerClass}`}
        >
          <Presentation className={iconClass} />
        </div>
      );
  }
};

export const DocumentThumbnail: React.FC<{
  type: OfficeModuleType;
  previewText: string;
}> = ({type, previewText}) => {
  return (
    <div className="relative flex h-24 w-full flex-col justify-between overflow-hidden rounded-lg border border-slate-200 bg-slate-50 p-3 dark:border-slate-800 dark:bg-slate-950">
      <div className="flex items-center justify-between gap-2">
        <ModuleIcon type={type} size="sm" />
        <span className="font-mono text-[11px] font-medium text-slate-500 dark:text-slate-400">
          {type.toUpperCase()}
        </span>
      </div>
      <p className="line-clamp-2 text-[11px] leading-relaxed text-slate-500 dark:text-slate-400">
        {previewText || 'Empty document'}
      </p>
    </div>
  );
};
