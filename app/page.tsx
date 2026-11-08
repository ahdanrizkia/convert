'use client';

import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  CATEGORIES,
  DEFAULT_TTL_MS,
  FORMAT_CATEGORY_MAP,
  MAX_FILE_SIZE_BYTES,
  formatBytes,
  getTargetFormats,
  normalizeFormat,
  type CategoryId,
} from '@/lib/converters/definitions';

type RowStatus =
  | 'ready'
  | 'converting'
  | 'done'
  | 'downloading'
  | 'downloaded_deleted'
  | 'purged'
  | 'error';

interface FileQueueItem {
  id: string;
  file: File | null;
  fileName: string;
  originalSize: number;
  convertedSize: number | null;
  sourceFormat: string;
  targetFormat: string;
  availableTargets: readonly string[];
  status: RowStatus;
  jobId: string | null;
  outputFileName: string | null;
  errorMessage: string | null;
  startTime: number | null;
  elapsedFinal: string | null;
}

function getStatusBadgeConfig(
  row: FileQueueItem,
  now: number
): { label: string; className: string } {
  switch (row.status) {
    case 'ready':
      return {
        label: 'Ready',
        className: 'border-neutral-400 bg-neutral-100 text-neutral-800',
      };
    case 'converting': {
      const elapsedSec = Math.max(0, (now - (row.startTime || now)) / 1000);
      let microStep = 'Uploading to RAM/tmp...';
      if (elapsedSec >= 2.5) {
        microStep = 'Finalizing output file...';
      } else if (elapsedSec >= 0.8) {
        microStep = `Processing ${row.sourceFormat.toUpperCase()} to ${row.targetFormat.toUpperCase()}...`;
      }
      return {
        label: `${microStep} (${elapsedSec.toFixed(1)}s)`,
        className: 'border-amber-700 bg-amber-50 text-amber-900',
      };
    }
    case 'done':
      return {
        label: row.elapsedFinal ? `Done in ${row.elapsedFinal}` : 'Done',
        className: 'border-emerald-700 bg-emerald-50 text-emerald-900',
      };
    case 'downloading':
      return {
        label: 'Downloading...',
        className: 'border-neutral-900 bg-neutral-200 text-neutral-950',
      };
    case 'downloaded_deleted':
      return {
        label: 'Downloaded & Auto-Deleted',
        className: 'border-neutral-500 bg-neutral-200 text-neutral-700',
      };
    case 'purged':
      return {
        label: 'Purged from server',
        className: 'border-neutral-500 bg-neutral-200 text-neutral-700',
      };
    case 'error':
      return {
        label: row.errorMessage || 'Error',
        className: 'border-red-700 bg-red-50 text-red-900',
      };
  }
}

export default function EphemeralConverterPage() {
  const [activeTab, setActiveTab] = useState<CategoryId>('images');
  const [items, setItems] = useState<FileQueueItem[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [now, setNow] = useState<number>(Date.now());

  const hasConverting = items.some((i) => i.status === 'converting');

  useEffect(() => {
    if (!hasConverting) return;
    const interval = setInterval(() => {
      setNow(Date.now());
    }, 100);
    return () => clearInterval(interval);
  }, [hasConverting]);

  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const ttlTimersRef = useRef<Map<string, ReturnType<typeof setTimeout>>>(new Map());

  const activeCategory =
    CATEGORIES.find((cat) => cat.id === activeTab) ?? CATEGORIES[0];

  // Clean up any client-side TTL sync timers on unmount
  useEffect(() => {
    const timersMap = ttlTimersRef.current;
    return () => {
      for (const timer of timersMap.values()) {
        clearTimeout(timer);
      }
      timersMap.clear();
    };
  }, []);

  const clearClientTtlTimer = useCallback((rowId: string) => {
    const existing = ttlTimersRef.current.get(rowId);
    if (existing) {
      clearTimeout(existing);
      ttlTimersRef.current.delete(rowId);
    }
  }, []);

  const startClientTtlTimer = useCallback(
    (rowId: string) => {
      clearClientTtlTimer(rowId);
      const timer = setTimeout(() => {
        ttlTimersRef.current.delete(rowId);
        setItems((prev) =>
          prev.map((row) =>
            row.id === rowId && row.status === 'done'
              ? { ...row, status: 'purged', jobId: null }
              : row
          )
        );
      }, DEFAULT_TTL_MS);
      ttlTimersRef.current.set(rowId, timer);
    },
    [clearClientTtlTimer]
  );

  const enqueueFiles = useCallback((fileList: FileList | File[]) => {
    const incoming = Array.from(fileList);
    if (incoming.length === 0) return;

    const newRows: FileQueueItem[] = incoming.map((file) => {
      const id =
        typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
          ? crypto.randomUUID()
          : `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;

      const hasDot = file.name.includes('.');
      const sourceFormat = hasDot ? normalizeFormat(file.name) : '';
      const availableTargets = sourceFormat ? getTargetFormats(sourceFormat) : [];

      // Strict frontend 50 MB check BEFORE network upload
      if (file.size > MAX_FILE_SIZE_BYTES) {
        return {
          id,
          file: null,
          fileName: file.name,
          originalSize: file.size,
          convertedSize: null,
          sourceFormat: sourceFormat || 'FILE',
          targetFormat: '—',
          availableTargets: [],
          status: 'error',
          jobId: null,
          outputFileName: null,
          errorMessage: 'Exceeds 50 MB limit',
          startTime: null,
          elapsedFinal: null,
        };
      }

      if (file.size === 0) {
        return {
          id,
          file: null,
          fileName: file.name,
          originalSize: 0,
          convertedSize: null,
          sourceFormat: sourceFormat || 'FILE',
          targetFormat: '—',
          availableTargets: [],
          status: 'error',
          jobId: null,
          outputFileName: null,
          errorMessage: 'Empty file (0 B)',
          startTime: null,
          elapsedFinal: null,
        };
      }

      if (availableTargets.length === 0) {
        return {
          id,
          file: null,
          fileName: file.name,
          originalSize: file.size,
          convertedSize: null,
          sourceFormat: sourceFormat || 'UNKNOWN',
          targetFormat: '—',
          availableTargets: [],
          status: 'error',
          jobId: null,
          outputFileName: null,
          errorMessage: 'Unsupported format',
          startTime: null,
          elapsedFinal: null,
        };
      }

      const defaultTarget =
        activeTab === 'documents' && availableTargets.includes('pdf')
          ? 'pdf'
          : availableTargets[0];

      return {
        id,
        file,
        fileName: file.name,
        originalSize: file.size,
        convertedSize: null,
        sourceFormat,
        targetFormat: defaultTarget,
        availableTargets,
        status: 'ready',
        jobId: null,
        outputFileName: null,
        errorMessage: null,
        startTime: null,
        elapsedFinal: null,
      };
    });

    // Keep user on [Documents] when queuing PDF-compatible inputs; otherwise switch to detected category
    const firstValid = newRows.find((r) => r.status === 'ready');
    if (firstValid) {
      const staysInDocuments =
        activeTab === 'documents' && firstValid.availableTargets.includes('pdf');
      if (!staysInDocuments) {
        const detectedCategory = FORMAT_CATEGORY_MAP[firstValid.sourceFormat];
        if (detectedCategory) {
          setActiveTab(detectedCategory);
        }
      }
    }

    setItems((prev) => [...newRows, ...prev]);
  }, [activeTab]);

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      enqueueFiles(e.target.files);
      e.target.value = '';
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    if (!activeCategory.active) return;
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      enqueueFiles(e.dataTransfer.files);
    }
  };

  const handleTargetChange = (rowId: string, nextTarget: string) => {
    setItems((prev) =>
      prev.map((row) =>
        row.id === rowId && row.status === 'ready'
          ? { ...row, targetFormat: nextTarget }
          : row
      )
    );
  };

  const handleRemoveRow = (rowId: string) => {
    clearClientTtlTimer(rowId);
    setItems((prev) => prev.filter((row) => row.id !== rowId));
  };

  const handleClearAll = () => {
    for (const timer of ttlTimersRef.current.values()) {
      clearTimeout(timer);
    }
    ttlTimersRef.current.clear();
    setItems([]);
  };

  const handleConvertRow = async (rowId: string) => {
    const targetRow = items.find((r) => r.id === rowId);
    if (!targetRow || targetRow.status !== 'ready' || !targetRow.file) return;

    const startMs = Date.now();
    setItems((prev) =>
      prev.map((r) =>
        r.id === rowId
          ? {
              ...r,
              status: 'converting',
              errorMessage: null,
              startTime: startMs,
              elapsedFinal: null,
            }
          : r
      )
    );

    try {
      const formData = new FormData();
      formData.append('file', targetRow.file);
      formData.append('targetFormat', targetRow.targetFormat);

      const response = await fetch('/api/convert', {
        method: 'POST',
        body: formData,
      });

      const payload = await response.json().catch(() => ({}));

      if (!response.ok) {
        throw new Error(payload?.error || `Conversion failed (${response.status})`);
      }

      const elapsedSec = ((Date.now() - startMs) / 1000).toFixed(1) + 's';

      setItems((prev) =>
        prev.map((r) =>
          r.id === rowId
            ? {
                ...r,
                // Release local File reference after upload to keep browser memory flat
                file: null,
                status: 'done',
                jobId: payload.jobId,
                outputFileName: payload.outputFileName,
                convertedSize:
                  typeof payload.convertedSize === 'number'
                    ? payload.convertedSize
                    : null,
                errorMessage: null,
                elapsedFinal: elapsedSec,
              }
            : r
        )
      );

      // Sync client badge to server's 3-minute abandonment TTL timer
      startClientTtlTimer(rowId);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Conversion failed';
      setItems((prev) =>
        prev.map((r) =>
          r.id === rowId
            ? {
                ...r,
                file: null,
                status: 'error',
                errorMessage: msg,
                elapsedFinal: null,
              }
            : r
        )
      );
    }
  };

  const handleConvertAllReady = async () => {
    const readyIds = items.filter((r) => r.status === 'ready').map((r) => r.id);
    // Process sequentially to avoid saturating low-bandwidth Wi-Fi connections
    for (const id of readyIds) {
      await handleConvertRow(id);
    }
  };

  const handleDownloadRow = async (rowId: string) => {
    const targetRow = items.find((r) => r.id === rowId);
    if (!targetRow || targetRow.status !== 'done' || !targetRow.jobId) return;

    clearClientTtlTimer(rowId);

    setItems((prev) =>
      prev.map((r) => (r.id === rowId ? { ...r, status: 'downloading' } : r))
    );

    try {
      const encodedName = encodeURIComponent(
        targetRow.outputFileName || `converted.${targetRow.targetFormat}`
      );
      const response = await fetch(
        `/api/download/${encodeURIComponent(targetRow.jobId)}?filename=${encodedName}`,
        {
          method: 'GET',
          cache: 'no-store',
        }
      );

      if (response.status === 410 || response.status === 404) {
        setItems((prev) =>
          prev.map((r) =>
            r.id === rowId ? { ...r, status: 'purged', jobId: null } : r
          )
        );
        return;
      }

      if (!response.ok) {
        throw new Error(`Download failed (${response.status})`);
      }

      // Stream full attachment body; server immediately unlinks `/tmp` output as stream ends
      const blob = await response.blob();
      const downloadUrl = URL.createObjectURL(blob);

      const anchor = document.createElement('a');
      anchor.href = downloadUrl;
      anchor.download =
        targetRow.outputFileName || `converted.${targetRow.targetFormat}`;
      anchor.style.display = 'none';
      document.body.appendChild(anchor);
      anchor.click();
      document.body.removeChild(anchor);

      // Revoke temporary blob URL immediately after browser dispatches save action
      setTimeout(() => {
        URL.revokeObjectURL(downloadUrl);
      }, 1000);

      // Update UI row status to "Downloaded & Auto-Deleted" and disable Download button
      setItems((prev) =>
        prev.map((r) =>
          r.id === rowId
            ? {
                ...r,
                status: 'downloaded_deleted',
                jobId: null,
              }
            : r
        )
      );
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Download interrupted';
      setItems((prev) =>
        prev.map((r) =>
          r.id === rowId
            ? {
                ...r,
                status: 'purged',
                jobId: null,
                errorMessage: msg,
              }
            : r
        )
      );
    }
  };

  const readyCount = items.filter((r) => r.status === 'ready').length;

  return (
    <main className="mx-auto max-w-4xl px-4 py-6 sm:px-6 sm:py-10">
      {/* Minimalist Utilitarian Header */}
      <header className="border border-neutral-900 bg-white p-4 sm:p-5">
        <div className="flex flex-wrap items-baseline justify-between gap-2">
          <div className="flex items-center gap-2.5">
            <span className="inline-block border border-neutral-950 bg-neutral-950 px-2 py-0.5 font-mono text-xs font-bold tracking-tight text-white">
              /tmp/conv
            </span>
            <h1 className="text-lg font-bold tracking-tight text-neutral-950 sm:text-xl">
              Stateless Ephemeral File Converter
            </h1>
          </div>
          <span className="font-mono text-xs text-neutral-600">
            MAX 50 MB · ZERO DB · ZERO LOGS
          </span>
        </div>

        <p className="mt-2 text-xs leading-relaxed text-neutral-700 sm:text-sm">
          Files convert inside isolated <code className="font-mono">/tmp</code> storage.
          Original uploads are unlinked immediately after conversion; converted outputs
          self-destruct the instant your download finishes (or after a 3-minute TTL).
        </p>

        <p className="mt-2.5 border-t border-neutral-200 pt-2 text-xs leading-relaxed text-neutral-800 sm:text-sm">
          <span className="font-semibold text-neutral-950">Petunjuk:</span> Cukup unggah atau drop file Anda (misal pilih tab <strong>[Documents]</strong> untuk <code className="font-mono font-semibold text-neutral-900">.docx</code>). Format tujuan konversi seperti <span className="font-mono font-semibold text-neutral-900">DOCX &rarr; PDF</span> dapat langsung dipilih pada daftar antrean di bawah.
        </p>
      </header>

      {/* Navigation Tabs for Phased Rollout */}
      <nav
        aria-label="Converter Categories"
        className="mt-4 grid grid-cols-2 gap-2 sm:grid-cols-4"
      >
        {CATEGORIES.map((category) => {
          const isSelected = category.id === activeTab;
          return (
            <button
              key={category.id}
              type="button"
              onClick={() => setActiveTab(category.id)}
              className={`flex flex-col items-start justify-between border px-3 py-2.5 text-left ${
                isSelected
                  ? 'border-neutral-950 bg-neutral-950 text-white'
                  : 'border-neutral-300 bg-white text-neutral-900 hover:border-neutral-900'
              }`}
            >
              <span className="text-xs font-bold sm:text-sm">[{category.label}]</span>
              <span
                className={`mt-1 font-mono text-[11px] ${
                  isSelected ? 'text-neutral-300' : 'text-neutral-600'
                }`}
              >
                {category.active
                  ? category.formats.map((f) => f.toUpperCase()).join(' · ')
                  : category.badge}
              </span>
            </button>
          );
        })}
      </nav>

      {/* Dropzone or Modular Phase Placeholder */}
      <section className="mt-4 border border-neutral-900 bg-white p-4 sm:p-6">
        {activeCategory.active ? (
          <div
            onDragOver={(e) => {
              e.preventDefault();
              e.stopPropagation();
              if (!isDragging) setIsDragging(true);
            }}
            onDragLeave={(e) => {
              e.preventDefault();
              e.stopPropagation();
              setIsDragging(false);
            }}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                fileInputRef.current?.click();
              }
            }}
            className={`cursor-pointer border border-dashed p-6 text-center sm:p-8 ${
              isDragging
                ? 'border-neutral-950 bg-neutral-100'
                : 'border-neutral-400 bg-neutral-50 hover:border-neutral-900'
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              multiple
              accept={activeCategory.accept}
              onChange={handleFileInputChange}
              className="hidden"
            />

            <div className="mx-auto flex max-w-md flex-col items-center gap-2">
              {/* Minimal 16x16 inline stroke SVG */}
              <svg
                width="20"
                height="20"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="square"
                strokeLinejoin="miter"
                aria-hidden="true"
                className="text-neutral-900"
              >
                <path d="M12 16V4" />
                <path d="M7 9l5-5 5 5" />
                <path d="M4 20h16" />
              </svg>

              <p className="text-sm font-semibold text-neutral-950">
                Select or drop {activeCategory.label} files to convert
              </p>

              <p className="font-mono text-xs text-neutral-600">
                Supported:{' '}
                {activeCategory.formats.map((f) => f.toUpperCase()).join(', ')} · Max 50
                MB/file · Zero DOM Previews
              </p>

              <span className="mt-1 inline-block border border-neutral-900 bg-white px-3 py-1 font-mono text-xs font-semibold text-neutral-950">
                BROWSE FILES
              </span>
            </div>
          </div>
        ) : (
          <div className="border border-neutral-300 bg-neutral-50 p-5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-sm font-bold text-neutral-950">
                [{activeCategory.label}] Engine Slot
              </h2>
              <span className="border border-neutral-900 bg-white px-2 py-0.5 font-mono text-xs font-semibold text-neutral-900">
                {activeCategory.badge}
              </span>
            </div>
            <p className="mt-2 text-xs text-neutral-700 sm:text-sm">
              {activeCategory.description} Wired into{' '}
              <code className="font-mono">/lib/converters/registry.ts</code> so handlers
              plug in without altering UI or <code className="font-mono">/tmp</code>{' '}
              auto-delete hooks.
            </p>
            {activeCategory.plannedRoutes && (
              <div className="mt-3 flex flex-wrap gap-1.5">
                {activeCategory.plannedRoutes.map((route) => (
                  <span
                    key={route}
                    className="border border-neutral-300 bg-white px-2 py-0.5 font-mono text-xs text-neutral-700"
                  >
                    {route}
                  </span>
                ))}
              </div>
            )}
          </div>
        )}
      </section>

      {/* Ephemeral Queue List (Zero Previews: Text & Monospace Badges Only) */}
      <section className="mt-4 border border-neutral-900 bg-white">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-neutral-900 bg-neutral-100 px-4 py-2.5">
          <div className="flex items-center gap-2">
            <h2 className="font-mono text-xs font-bold uppercase tracking-wider text-neutral-900">
              Conversion Queue ({items.length})
            </h2>
          </div>

          {items.length > 0 && (
            <div className="flex items-center gap-2">
              {readyCount > 1 && (
                <button
                  type="button"
                  onClick={handleConvertAllReady}
                  className="border border-neutral-950 bg-neutral-950 px-2.5 py-1 font-mono text-xs font-semibold text-white hover:bg-neutral-800"
                >
                  Convert All ({readyCount})
                </button>
              )}
              <button
                type="button"
                onClick={handleClearAll}
                className="border border-neutral-400 bg-white px-2.5 py-1 font-mono text-xs text-neutral-800 hover:border-neutral-950"
              >
                Clear List
              </button>
            </div>
          )}
        </div>

        {items.length === 0 ? (
          <div className="px-4 py-8 text-center font-mono text-xs text-neutral-500">
            No files queued. Select Images, Office Documents (PDF, DOCX, PPTX, XLSX), Data, or Audio/Video files above.
          </div>
        ) : (
          <ul className="divide-y divide-neutral-200">
            {items.map((row) => {
              const statusBadge = getStatusBadgeConfig(row, now);

              return (
                <li
                  key={row.id}
                  className="relative flex flex-col gap-2.5 px-4 py-3 sm:flex-row sm:items-center sm:justify-between sm:gap-3 overflow-hidden"
                >
                  {/* 1. File Name (cleanly truncated with ellipsis) + 2. Formatted File Size */}
                  <div className="flex min-w-0 flex-1 items-center gap-3">
                    <span
                      title={row.fileName}
                      className="truncate text-sm font-medium text-neutral-950"
                    >
                      {row.fileName}
                    </span>
                    <span className="shrink-0 font-mono text-xs text-neutral-600">
                      {row.convertedSize !== null
                        ? `${formatBytes(row.originalSize)} -> ${formatBytes(row.convertedSize)}`
                        : formatBytes(row.originalSize)}
                    </span>
                  </div>

                  {/* Right Group: 3. Route Badge, 4. Status Badge, 5. Action Button */}
                  <div className="flex flex-wrap items-center justify-between gap-2 sm:shrink-0 sm:justify-end">
                    {/* 3. Conversion Route Badge: [SOURCE] -> [TARGET Dropdown/Badge] */}
                    <div className="inline-flex items-center border border-neutral-300 bg-neutral-50 px-2 py-1 font-mono text-xs text-neutral-900">
                      <span className="font-semibold">
                        {row.sourceFormat.toUpperCase()}
                      </span>
                      <span className="mx-1.5 text-neutral-500">-&gt;</span>
                      {row.status === 'ready' && row.availableTargets.length > 0 ? (
                        <select
                          aria-label={`Target format for ${row.fileName}`}
                          value={row.targetFormat}
                          onChange={(e) => handleTargetChange(row.id, e.target.value)}
                          className="border border-neutral-400 bg-white px-1.5 py-0.5 font-mono text-xs font-bold text-neutral-950 focus:border-neutral-950 focus:outline-none"
                        >
                          {row.availableTargets.map((target) => (
                            <option key={target} value={target}>
                              {target.toUpperCase()}
                            </option>
                          ))}
                        </select>
                      ) : (
                        <span className="font-bold">
                          {row.targetFormat.toUpperCase()}
                        </span>
                      )}
                    </div>

                    {/* 4. Status Badge */}
                    <span
                      title={row.errorMessage ?? statusBadge.label}
                      className={`inline-flex max-w-[280px] items-center gap-1.5 truncate border px-2 py-1 font-mono text-xs font-medium ${statusBadge.className}`}
                    >
                      {row.status === 'converting' && (
                        <svg
                          className="h-3.5 w-3.5 shrink-0 animate-spin text-amber-900"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          aria-hidden="true"
                        >
                          <circle
                            className="opacity-25"
                            cx="12"
                            cy="12"
                            r="10"
                            stroke="currentColor"
                            strokeWidth="4"
                          />
                          <path
                            className="opacity-75"
                            fill="currentColor"
                            d="M4 12a8 8 0 018-8V4a4 4 0 00-4 4H4z"
                          />
                        </svg>
                      )}
                      <span className="truncate">{statusBadge.label}</span>
                    </span>

                    {/* 5. Action Button (Convert, Download, or Remove) */}
                    <div className="flex items-center gap-1.5">
                      {row.status === 'ready' && (
                        <>
                          <button
                            type="button"
                            onClick={() => handleConvertRow(row.id)}
                            className="border border-neutral-950 bg-neutral-950 px-3 py-1 font-mono text-xs font-semibold text-white hover:bg-neutral-800"
                          >
                            Convert
                          </button>
                          <button
                            type="button"
                            onClick={() => handleRemoveRow(row.id)}
                            aria-label={`Remove ${row.fileName}`}
                            className="border border-neutral-300 bg-white px-2 py-1 font-mono text-xs text-neutral-700 hover:border-neutral-900"
                          >
                            Remove
                          </button>
                        </>
                      )}

                      {row.status === 'converting' && (
                        <button
                          type="button"
                          disabled
                          className="cursor-not-allowed border border-neutral-300 bg-neutral-100 px-3 py-1 font-mono text-xs text-neutral-500"
                        >
                          Convert
                        </button>
                      )}

                      {row.status === 'done' && (
                        <button
                          type="button"
                          onClick={() => handleDownloadRow(row.id)}
                          className="border border-emerald-800 bg-emerald-700 px-3 py-1 font-mono text-xs font-semibold text-white hover:bg-emerald-800"
                        >
                          Download
                        </button>
                      )}

                      {row.status === 'downloading' && (
                        <button
                          type="button"
                          disabled
                          className="cursor-not-allowed border border-neutral-300 bg-neutral-100 px-3 py-1 font-mono text-xs text-neutral-500"
                        >
                          Download
                        </button>
                      )}

                      {(row.status === 'downloaded_deleted' ||
                        row.status === 'purged') && (
                        <button
                          type="button"
                          disabled
                          title="File has been permanently deleted from server /tmp storage"
                          className="cursor-not-allowed border border-neutral-300 bg-neutral-100 px-3 py-1 font-mono text-xs text-neutral-400 line-through"
                        >
                          Download
                        </button>
                      )}

                      {row.status === 'error' && (
                        <button
                          type="button"
                          onClick={() => handleRemoveRow(row.id)}
                          className="border border-neutral-900 bg-white px-3 py-1 font-mono text-xs font-semibold text-neutral-900 hover:bg-neutral-100"
                        >
                          Remove
                        </button>
                      )}
                    </div>
                  </div>

                  {/* 4. Thin Indeterminate Progress Bar */}
                  {row.status === 'converting' && (
                    <div className="absolute bottom-0 left-0 right-0 h-[2px] bg-amber-200 overflow-hidden">
                      <div className="h-full w-1/3 bg-amber-700 animate-indeterminate" />
                    </div>
                  )}
                </li>
              );
            })}
          </ul>
        )}
      </section>

      {/* Compact Ephemeral Lifecycle Guarantee Footer */}
      <footer className="mt-4 flex flex-wrap items-center justify-between gap-2 border border-neutral-300 bg-white px-4 py-2.5 font-mono text-[11px] text-neutral-600">
        <span>PIPELINE: UPLOAD -&gt; /tmp/UUID -&gt; UNLINK INPUT -&gt; STREAM &amp; UNLINK OUTPUT</span>
        <span>FALLBACK TTL: 180s</span>
      </footer>
    </main>
  );
}
