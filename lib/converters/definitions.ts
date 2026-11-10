export const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024; // Strict 50 MB limit
export const DEFAULT_TTL_MS = 180_000; // 3 minutes (180,000 ms) fallback abandonment timer

export type CategoryId = 'images' | 'data' | 'documents' | 'media';

export type ImageFormat =
  | 'png'
  | 'jpg'
  | 'jfif'
  | 'webp'
  | 'avif'
  | 'svg'
  | 'gif'
  | 'bmp'
  | 'ico'
  | 'tiff';

export type DataFormat = 'json' | 'csv' | 'txt';

export type DocumentFormat = 'pdf' | 'docx' | 'pptx' | 'xlsx' | 'md' | 'html' | 'epub';

export type MediaFormat =
  | 'mp4'
  | 'mp3'
  | 'wav'
  | 'ogg'
  | 'aac'
  | 'm4a'
  | 'mov'
  | 'webm'
  | 'mkv';

export const PDF_SOURCE_FORMATS = [
  'png',
  'jpg',
  'jfif',
  'webp',
  'avif',
  'svg',
  'gif',
  'tiff',
  'bmp',
  'ico',
  'txt',
  'docx',
  'pptx',
  'xlsx',
  'md',
  'html',
  'epub',
] as const;
export type PdfSourceFormat = (typeof PDF_SOURCE_FORMATS)[number];

export const MEDIA_SOURCE_FORMATS = [
  'mp4',
  'mp3',
  'wav',
  'ogg',
  'aac',
  'm4a',
  'mov',
  'webm',
  'mkv',
  'gif',
] as const;
export type MediaSourceFormat = (typeof MEDIA_SOURCE_FORMATS)[number];

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
    label: 'Images & Vector',
    phase: 1,
    active: true,
    description: 'Lossless & lossy raster & vector conversion: PNG, JPG, JFIF, WEBP, AVIF, SVG, GIF, BMP, ICO, TIFF.',
    formats: ['png', 'jpg', 'jfif', 'webp', 'avif', 'svg', 'gif', 'bmp', 'ico', 'tiff'],
    accept:
      '.png,.jpg,.jpeg,.jfif,.webp,.avif,.svg,.gif,.bmp,.ico,.tiff,.tif,image/png,image/jpeg,image/pjpeg,image/webp,image/avif,image/svg+xml,image/gif,image/bmp,image/x-icon,image/tiff',
  },
  {
    id: 'documents',
    label: 'Documents',
    phase: 2,
    active: true,
    description: 'Stateless Office & Document suite: PDF, Word (DOCX), Slides (PPTX), Sheets (XLSX), Markdown (MD), HTML, EPUB.',
    formats: ['pdf', 'docx', 'pptx', 'xlsx', 'md', 'html', 'epub'],
    accept:
      '.pdf,.docx,.pptx,.xlsx,.md,.markdown,.html,.htm,.epub,.txt,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/markdown,text/html,application/epub+zip,text/plain',
  },
  {
    id: 'data',
    label: 'Data / Text',
    phase: 1,
    active: true,
    description: 'Streaming structured data, plain-text, and Excel transformations: JSON, CSV, TXT, XLSX.',
    formats: ['json', 'csv', 'txt', 'xlsx'],
    accept:
      '.json,.csv,.txt,.xlsx,application/json,text/csv,text/plain,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
  {
    id: 'media',
    label: 'Audio / Video',
    phase: 3,
    active: true,
    description: 'Audio extraction & transcoding: MP4, MP3, WAV, OGG, AAC, M4A, MOV, WEBM, MKV, GIF.',
    formats: ['mp4', 'mp3', 'wav', 'ogg', 'aac', 'm4a', 'mov', 'webm', 'mkv'],
    accept:
      '.mp4,.mp3,.wav,.ogg,.aac,.m4a,.mov,.webm,.mkv,.gif,video/mp4,audio/mpeg,audio/wav,audio/x-wav,audio/ogg,audio/aac,audio/mp4,video/quicktime,video/webm,video/x-matroska,image/gif',
  },
] as const;

export const FORMAT_ROUTES: Readonly<Record<string, readonly string[]>> = {
  // 1. Images & Vector
  png: ['webp', 'jpg', 'jfif', 'avif', 'ico', 'bmp', 'tiff', 'pdf', 'docx', 'pptx'],
  jpg: ['webp', 'png', 'jfif', 'avif', 'ico', 'bmp', 'tiff', 'pdf', 'docx', 'pptx'],
  jfif: ['jpg', 'png', 'webp', 'avif', 'ico', 'bmp', 'tiff', 'pdf', 'docx', 'pptx'],
  webp: ['png', 'jpg', 'jfif', 'avif', 'ico', 'bmp', 'tiff', 'pdf', 'docx', 'pptx'],
  avif: ['webp', 'png', 'jpg', 'jfif', 'ico', 'bmp', 'tiff', 'pdf', 'docx', 'pptx'],
  svg: ['png', 'jpg', 'jfif', 'webp', 'avif', 'pdf'],
  gif: ['mp4', 'webp', 'png', 'jpg'],
  bmp: ['png', 'jpg', 'webp'],
  ico: ['png', 'jpg', 'webp'],
  tiff: ['png', 'jpg', 'jfif', 'webp', 'avif', 'pdf'],

  // 2. Data / Text
  json: ['csv', 'txt', 'xlsx'],
  csv: ['json', 'txt', 'xlsx'],
  txt: ['json', 'csv', 'pdf', 'docx', 'pptx'],

  // 3. Documents (Office, Markdown, HTML, EPUB)
  docx: ['pdf', 'pptx', 'txt', 'html'],
  pdf: ['docx', 'pptx', 'png', 'jpg', 'jfif', 'txt'],
  xlsx: ['pdf', 'pptx', 'csv', 'json'],
  pptx: ['pdf', 'png', 'jpg', 'txt'],
  md: ['html', 'pdf', 'docx', 'txt'],
  html: ['pdf', 'docx', 'txt'],
  epub: ['pdf', 'txt'],

  // 4. Audio & Video
  mp4: ['mp3', 'wav'],
  mp3: ['wav', 'mp4'],
  wav: ['mp3', 'mp4'],
  ogg: ['mp3', 'wav'],
  aac: ['mp3', 'wav'],
  m4a: ['mp3', 'wav'],
  mov: ['mp4', 'mp3', 'wav'],
  webm: ['mp4', 'mp3', 'wav'],
  mkv: ['mp4', 'mp3', 'wav'],
};

export const FORMAT_CATEGORY_MAP: Readonly<Record<string, CategoryId>> = {
  // Images
  png: 'images',
  jpg: 'images',
  jfif: 'images',
  webp: 'images',
  avif: 'images',
  svg: 'images',
  gif: 'images',
  bmp: 'images',
  ico: 'images',
  tiff: 'images',

  // Data
  json: 'data',
  csv: 'data',
  txt: 'data',

  // Documents
  docx: 'documents',
  pdf: 'documents',
  xlsx: 'documents',
  pptx: 'documents',
  md: 'documents',
  html: 'documents',
  epub: 'documents',

  // Media
  mp4: 'media',
  mp3: 'media',
  wav: 'media',
  ogg: 'media',
  aac: 'media',
  m4a: 'media',
  mov: 'media',
  webm: 'media',
  mkv: 'media',
};

export const MIME_TYPES: Readonly<Record<string, string>> = {
  // Images
  png: 'image/png',
  jpg: 'image/jpeg',
  jfif: 'image/jpeg',
  webp: 'image/webp',
  avif: 'image/avif',
  svg: 'image/svg+xml',
  gif: 'image/gif',
  bmp: 'image/bmp',
  ico: 'image/x-icon',
  tiff: 'image/tiff',

  // Data
  json: 'application/json; charset=utf-8',
  csv: 'text/csv; charset=utf-8',
  txt: 'text/plain; charset=utf-8',

  // Documents
  pdf: 'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  html: 'text/html; charset=utf-8',
  md: 'text/markdown; charset=utf-8',
  epub: 'application/epub+zip',

  // Media
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  mp4: 'video/mp4',
  ogg: 'audio/ogg',
  aac: 'audio/aac',
  m4a: 'audio/mp4',
  mov: 'video/quicktime',
  webm: 'video/webm',
  mkv: 'video/x-matroska',
};

export const ALLOWED_OUTPUT_EXTENSIONS = Object.keys(MIME_TYPES);

/**
 * Normalizes a filename or raw extension into a canonical lowercase format key.
 * Maps `.jpeg` -> `jpg`, `.tif` -> `tiff`, `.markdown` -> `md`, `.htm` -> `html`.
 */
export function normalizeFormat(fileNameOrExt: string): string {
  const trimmed = fileNameOrExt.trim().toLowerCase();
  const ext = trimmed.includes('.')
    ? trimmed.slice(trimmed.lastIndexOf('.') + 1)
    : trimmed;

  if (ext === 'jpeg') return 'jpg';
  if (ext === 'tif') return 'tiff';
  if (ext === 'markdown') return 'md';
  if (ext === 'htm') return 'html';
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
