export const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024; // Strict 50 MB limit
export const DEFAULT_TTL_MS = 180_000; // 3 minutes (180,000 ms) fallback abandonment timer

export type CategoryId = 'images' | 'data' | 'documents' | 'media';

export type ImageFormat = 'png' | 'jpg' | 'webp' | 'avif';
export type DataFormat = 'json' | 'csv' | 'txt';
export type DocumentFormat = 'pdf';
export type Phase1Format = ImageFormat | DataFormat;

export const PDF_SOURCE_FORMATS = ['png', 'jpg', 'webp', 'avif', 'txt'] as const;
export type PdfSourceFormat = (typeof PDF_SOURCE_FORMATS)[number];

export const MEDIA_SOURCE_FORMATS = ['mp4', 'wav'] as const;
export type MediaSourceFormat = (typeof MEDIA_SOURCE_FORMATS)[number];
export type MediaTargetFormat = 'mp3';

export interface CategoryDefinition {
  id: CategoryId;
  label: string;
  phase: 1 | 2 | 3;
  active: boolean;
  badge?: string;
  description: string;
  formats: readonly string[];
  accept: string;
  plannedRoutes?: readonly string[];
}

export const CATEGORIES: readonly CategoryDefinition[] = [
  {
    id: 'images',
    label: 'Images',
    phase: 1,
    active: true,
    description: 'Lossless & lossy raster conversion powered by libvips (sharp).',
    formats: ['png', 'jpg', 'webp', 'avif'],
    accept: '.png,.jpg,.jpeg,.webp,.avif,image/png,image/jpeg,image/webp,image/avif',
  },
  {
    id: 'data',
    label: 'Data/Text',
    phase: 1,
    active: true,
    description: 'Streaming structured data, plain-text, and Excel transformations.',
    formats: ['json', 'csv', 'txt', 'xlsx'],
    accept: '.json,.csv,.txt,.xlsx,application/json,text/csv,text/plain,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
  {
    id: 'documents',
    label: 'Documents',
    phase: 2,
    active: true,
    description: 'Stateless Office & Document suite: PDF, Word (DOCX), Slides (PPTX), Sheets (XLSX).',
    formats: ['pdf', 'docx', 'pptx', 'xlsx'],
    accept: '.pdf,.docx,.pptx,.xlsx,.txt,.png,.jpg,.jpeg,.webp,.avif,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,image/png,image/jpeg,text/plain',
  },
  {
    id: 'media',
    label: 'Audio / Video',
    phase: 3,
    active: true,
    description: 'Stateless MP4-to-MP3 & WAV-to-MP3 audio extraction via isolated FFmpeg.',
    formats: ['mp4 -> mp3', 'wav -> mp3'],
    accept: '.mp4,.wav,video/mp4,audio/mp4,audio/wav,audio/x-wav',
  },
] as const;

export const FORMAT_ROUTES: Readonly<Record<string, readonly string[]>> = {
  // Phase 1: Images (+ Phase 2: Image -> PDF, DOCX, PPTX)
  png: ['webp', 'jpg', 'avif', 'pdf', 'docx', 'pptx'],
  jpg: ['webp', 'png', 'avif', 'pdf', 'docx', 'pptx'],
  webp: ['png', 'jpg', 'avif', 'pdf', 'docx', 'pptx'],
  avif: ['webp', 'png', 'jpg', 'pdf', 'docx', 'pptx'],
  // Phase 1: Data / Text (+ XLSX)
  json: ['csv', 'txt', 'xlsx'],
  csv: ['json', 'txt', 'xlsx'],
  txt: ['json', 'csv', 'pdf'],
  // Phase 2: Documents (PDF, DOCX, PPTX, XLSX)
  docx: ['pdf', 'txt', 'html'],
  pdf: ['docx', 'png', 'jpg', 'txt'],
  xlsx: ['pdf', 'csv', 'json'],
  pptx: ['pdf', 'txt'],
  // Phase 3: Audio / Video
  mp4: ['mp3'],
  wav: ['mp3'],
};

export const FORMAT_CATEGORY_MAP: Readonly<Record<string, CategoryId>> = {
  png: 'images',
  jpg: 'images',
  webp: 'images',
  avif: 'images',
  json: 'data',
  csv: 'data',
  txt: 'data',
  docx: 'documents',
  pdf: 'documents',
  xlsx: 'documents',
  pptx: 'documents',
  mp4: 'media',
  wav: 'media',
};

export const MIME_TYPES: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  webp: 'image/webp',
  avif: 'image/avif',
  json: 'application/json; charset=utf-8',
  csv: 'text/csv; charset=utf-8',
  txt: 'text/plain; charset=utf-8',
  pdf: 'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  html: 'text/html; charset=utf-8',
  mp3: 'audio/mpeg',
};

export const ALLOWED_OUTPUT_EXTENSIONS = Object.keys(MIME_TYPES);

/**
 * Normalizes a filename or raw extension into a canonical lowercase format key.
 * Maps `.jpeg` -> `jpg`.
 */
export function normalizeFormat(fileNameOrExt: string): string {
  const trimmed = fileNameOrExt.trim().toLowerCase();
  const ext = trimmed.includes('.')
    ? trimmed.slice(trimmed.lastIndexOf('.') + 1)
    : trimmed;

  if (ext === 'jpeg') return 'jpg';
  return ext;
}

/**
 * Returns valid target formats for a given source extension.
 */
export function getTargetFormats(sourceFormat: string): readonly string[] {
  const normalized = normalizeFormat(sourceFormat);
  return FORMAT_ROUTES[normalized] ?? [];
}

/**
 * Produces a safe, clean download filename with the converted extension.
 */
export function buildOutputFileName(originalName: string, targetFormat: string): string {
  const cleanTarget = normalizeFormat(targetFormat);
  // Strip any directory components and control/header-unsafe characters
  const base = originalName
    .replace(/^.*[\\/]/, '')
    .replace(/[^\w.\- ()[\]]+/g, '_')
    .trim();

  const dotIdx = base.lastIndexOf('.');
  const stem = dotIdx > 0 ? base.slice(0, dotIdx) : base || 'converted';
  const truncatedStem = stem.slice(0, 120);
  return `${truncatedStem}.${cleanTarget}`;
}

/**
 * Formats byte counts cleanly into B, KB, or MB (e.g. `1.42 MB`, `38.4 KB`).
 */
export function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return '0 B';
  if (bytes < 1024) return `${bytes} B`;
  const kb = bytes / 1024;
  if (kb < 1024) {
    return `${kb < 10 ? kb.toFixed(2) : kb.toFixed(1)} KB`;
  }
  const mb = kb / 1024;
  return `${mb.toFixed(2)} MB`;
}
