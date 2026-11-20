import { createReadStream } from 'node:fs';
import fs from 'node:fs/promises';
import readline from 'node:readline';
import path from 'node:path';
import {
  PDFDocument,
  PageSizes,
  StandardFonts,
  rgb,
  type PDFFont,
  type PDFImage,
} from 'pdf-lib';
import sharp from 'sharp';
import { isImageFormat } from './imageConverter';
import {
  normalizeFormat,
  type PdfPageLayout,
  type PdfSourceFormat,
  PDF_SOURCE_FORMATS,
} from './definitions';

// Keep libvips cache disabled so temporary buffers/descriptors release immediately
sharp.cache(false);

const SUPPORTED_PDF_SOURCES = new Set<string>(PDF_SOURCE_FORMATS);

// ISO 32000-1 (PDF 1.7) maximum user-space page dimension in points (200 x 200 in)
const MAX_PDF_PAGE_DIMENSION_PT = 14_400;

// Typography & layout constants for TXT -> PDF (A4 portrait)
const PAGE_MARGIN_X = 50;
const PAGE_MARGIN_TOP = 50;
const PAGE_MARGIN_BOTTOM = 50;
const TEXT_FONT_SIZE = 10.5;
const TEXT_LINE_HEIGHT = 15;
const TEXT_COLOR = rgb(0.08, 0.08, 0.08);

export function isPdfSourceFormat(fmt: string): fmt is PdfSourceFormat {
  return SUPPORTED_PDF_SOURCES.has(fmt);
}

/**
 * Builds a fast character-level sanitizer ensuring every string passed to pdf-lib's
 * standard WinAnsi font encodes cleanly without throwing on tabs, emojis, or symbols.
 */
function buildLineSanitizer(font: PDFFont): (raw: string) => string {
  const supported = new Set<number>(font.getCharacterSet());
  const replacementCache = new Map<string, string>();

  const mapUnsupportedChar = (ch: string): string => {
    const cached = replacementCache.get(ch);
    if (cached !== undefined) return cached;

    const decomposed = ch.normalize('NFKD');
    let fallback = '';
    for (const part of decomposed) {
      const cp = part.codePointAt(0);
      if (cp !== undefined && cp >= 0x20 && cp !== 0x7f && supported.has(cp)) {
        fallback += part;
      }
    }

    const resolved = fallback.length > 0 ? fallback : '?';
    replacementCache.set(ch, resolved);
    return resolved;
  };

  return (raw: string): string => {
    const preNormalized = raw
      .replace(/\t/g, '    ')
      .replace(/\u00a0/g, ' ')
      .replace(/[\u200b-\u200d\ufeff]/g, '')
      .replace(/\u2192/g, '->')
      .replace(/\u2190/g, '<-')
      .replace(/\u2194/g, '<->')
      .replace(/\u21d2/g, '=>')
      .replace(/[\u2713\u2714]/g, '[x]')
      .replace(/[\u2717\u2718]/g, '[ ]')
      .replace(/\u2212/g, '-');

    let out = '';
    for (const ch of preNormalized) {
      const cp = ch.codePointAt(0);
      if (cp === undefined || cp < 0x20 || cp === 0x7f) {
        continue;
      }
      if (supported.has(cp)) {
        out += ch;
      } else {
        out += mapUnsupportedChar(ch);
      }
    }
    return out;
  };
}

/**
 * Wraps a single sanitized text line to fit within `maxWidth` points,
 * preserving blank lines, leading indentation, and breaking oversized unbroken tokens.
 */
function wrapTextLine(
  line: string,
  font: PDFFont,
  fontSize: number,
  maxWidth: number
): string[] {
  if (line.length === 0) {
    return [''];
  }

  if (font.widthOfTextAtSize(line, fontSize) <= maxWidth) {
    return [line];
  }

  const wrapped: string[] = [];
  const tokens = line.match(/\S+|\s+/g) ?? [line];
  let current = '';

  const pushLongTokenChunks = (token: string): void => {
    let chunk = '';
    for (const ch of token) {
      const next = chunk + ch;
      if (chunk.length > 0 && font.widthOfTextAtSize(next, fontSize) > maxWidth) {
        wrapped.push(chunk);
        chunk = ch;
      } else {
        chunk = next;
      }
    }
    current = chunk;
  };

  for (const token of tokens) {
    if (current.length === 0) {
      if (font.widthOfTextAtSize(token, fontSize) <= maxWidth) {
        current = token;
      } else if (/^\s+$/.test(token)) {
        current = '';
      } else {
        pushLongTokenChunks(token);
      }
      continue;
    }

    const candidate = current + token;
    if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth) {
      current = candidate;
      continue;
    }

    wrapped.push(current.trimEnd());
    current = '';

    if (/^\s+$/.test(token)) {
      continue;
    }

    if (font.widthOfTextAtSize(token, fontSize) <= maxWidth) {
      current = token;
    } else {
      pushLongTokenChunks(token);
    }
  }

  if (current.length > 0) {
    wrapped.push(current.trimEnd());
  }

  return wrapped.length > 0 ? wrapped : [''];
}

/**
 * Embeds an image file into a `PDFDocument` with maximum visual fidelity:
 * - Upright sRGB/Grayscale JPEG (`jpg`, `jfif`) streams are embedded bit-for-bit without re-encoding (0% generation loss).
 * - JPEGs requiring EXIF orientation rotation or CMYK->sRGB color-profile conversion are encoded at Q98 with 4:4:4 chroma (zero color subsampling loss).
 * - PNG, SVG (300 DPI vector rasterization), WebP, AVIF, TIFF, BMP, ICO, and GIF are EXIF-rotated, normalized to 8-bit non-interlaced sRGB/RGBA lossless PNG, and embedded via FlateDecode to preserve 100% sharp text/edges and alpha transparency without `upng-js` interlace/16-bit/palette artifacts.
 */
async function embedHighAccuracyImage(
  pdfDoc: PDFDocument,
  inputPath: string,
  rawSourceFormat: string
): Promise<PDFImage> {
  // Read into buffer immediately so the `/tmp` input file descriptor closes right away
  let inputBuffer: Buffer = await fs.readFile(inputPath);
  let sourceFormat = normalizeFormat(rawSourceFormat);

  // Check if BMP, ICO, or an unreadable AVIF (e.g. 10-bit HDR / AV1 profile) requires FFmpeg lossless PNG decoding
  let needsFfmpegTranscode = sourceFormat === 'bmp' || sourceFormat === 'ico';
  if (!needsFfmpegTranscode && sourceFormat === 'avif') {
    try {
      await sharp(inputBuffer, { failOn: 'none' }).metadata();
    } catch {
      needsFfmpegTranscode = true;
    }
  }

  if (needsFfmpegTranscode) {
    const { spawn } = await import('node:child_process');
    const ffmpegStatic = (await import('ffmpeg-static')).default;
    const exeName = process.platform === 'win32' ? 'ffmpeg.exe' : 'ffmpeg';
    const ffmpegBin =
      process.env.FFMPEG_PATH?.trim() ||
      (typeof ffmpegStatic === 'string'
        ? ffmpegStatic
        : path.join(process.cwd(), 'node_modules', 'ffmpeg-static', exeName));

    const pngChunks: Buffer[] = [];
    await new Promise<void>((resolve, reject) => {
      const child = spawn(
        ffmpegBin,
        [
          '-hide_banner',
          '-loglevel',
          'error',
          '-y',
          '-i',
          inputPath,
          '-f',
          'image2pipe',
          '-vcodec',
          'png',
          '-pix_fmt',
          'rgba',
          '-',
        ],
        {
          stdio: ['ignore', 'pipe', 'pipe'],
          windowsHide: true,
        }
      );
      child.stdout.on('data', (d: Buffer) => pngChunks.push(d));
      child.once('error', reject);
      child.once('close', (code) => {
        if (code === 0) resolve();
        else reject(new Error(`Failed to decode ${sourceFormat.toUpperCase()} for PDF`));
      });
    });
    inputBuffer = Buffer.concat(pngChunks);
    sourceFormat = 'png';
  }

  const sharpOptions =
    sourceFormat === 'svg'
      ? { density: 300, failOn: 'none' as const }
      : { failOn: 'none' as const };
  const metadata = await sharp(inputBuffer, sharpOptions).metadata();

  const needsExifRotation =
    typeof metadata.orientation === 'number' && metadata.orientation > 1;
  const isNonSrgb =
    metadata.space !== undefined &&
    metadata.space !== 'srgb' &&
    metadata.space !== 'b-w';
  const isJpeg = sourceFormat === 'jpg' || sourceFormat === 'jfif';

  if (isJpeg) {
    if (!needsExifRotation && !isNonSrgb) {
      try {
        // 100% lossless direct DCT stream embedding
        return await pdfDoc.embedJpg(inputBuffer);
      } catch {
        // Fallback if JPEG uses arithmetic coding or non-standard markers
      }
    }
    const normalizedJpg = await sharp(inputBuffer, sharpOptions)
      .rotate()
      .toColorspace('srgb')
      .flatten({ background: '#ffffff' })
      .jpeg({ quality: 98, mozjpeg: true, chromaSubsampling: '4:4:4' })
      .toBuffer();
    return await pdfDoc.embedJpg(normalizedJpg);
  }

  // For PNG, SVG, WebP, AVIF, TIFF, BMP, ICO, GIF:
  // Embed as lossless sRGB/RGBA non-interlaced PNG to guarantee zero DCT compression artifacts
  // and full compatibility with pdf-lib's PNG parser across 16-bit, Adam7 interlaced, and palette images.
  const pixelCount = (metadata.width ?? 0) * (metadata.height ?? 0);
  const isUltraHugeOpaquePhoto =
    pixelCount > 25_000_000 && !metadata.hasAlpha && sourceFormat !== 'svg';

  try {
    const pipeline = sharp(inputBuffer, sharpOptions).rotate().toColorspace('srgb');
    if (isUltraHugeOpaquePhoto) {
      const highFidelityJpg = await pipeline
        .flatten({ background: '#ffffff' })
        .jpeg({ quality: 98, mozjpeg: true, chromaSubsampling: '4:4:4' })
        .toBuffer();
      return await pdfDoc.embedJpg(highFidelityJpg);
    }

    const losslessPng = await pipeline
      .png({ compressionLevel: 8, adaptiveFiltering: true })
      .toBuffer();
    return await pdfDoc.embedPng(losslessPng);
  } catch (err) {
    if (sourceFormat === 'png') {
      return await pdfDoc.embedPng(inputBuffer);
    }
    throw err;
  }
}

/**
 * Adds a page containing `embeddedImage` to `pdfDoc` according to the requested `layout` mode.
 * Preserves 100% of the embedded image's native pixel resolution inside the PDF XObject stream.
 */
function addImagePageToPdf(
  pdfDoc: PDFDocument,
  embeddedImage: PDFImage,
  layout: PdfPageLayout
): void {
  const imgW = embeddedImage.width;
  const imgH = embeddedImage.height;

  if (imgW <= 0 || imgH <= 0) {
    throw new Error('Invalid image dimensions for PDF generation.');
  }

  if (layout === 'a4-fit') {
    const isLandscape = imgW > imgH;
    const [a4Short, a4Long] = PageSizes.A4; // 595.28 x 841.89 pt
    const pageW = isLandscape ? a4Long : a4Short;
    const pageH = isLandscape ? a4Short : a4Long;
    const margin = 18;
    const availW = Math.max(1, pageW - margin * 2);
    const availH = Math.max(1, pageH - margin * 2);
    const scale = Math.min(availW / imgW, availH / imgH);
    const drawW = imgW * scale;
    const drawH = imgH * scale;
    const x = (pageW - drawW) / 2;
    const y = (pageH - drawH) / 2;

    const page = pdfDoc.addPage([pageW, pageH]);
    page.drawImage(embeddedImage, {
      x,
      y,
      width: drawW,
      height: drawH,
    });
    return;
  }

  if (layout === 'fit-image') {
    // Full-bleed (zero margin, exact image aspect ratio) normalized to a uniform A4 reference width (595.28 pt)
    // so multi-page PDFs with varying source pixel resolutions scroll at a consistent width in all PDF viewers
    // while retaining 100% of their original embedded pixel resolution.
    const refWidth = PageSizes.A4[0]; // 595.28 pt (8.27 in)
    let pageW = refWidth;
    let pageH = (refWidth * imgH) / imgW;

    if (pageH > MAX_PDF_PAGE_DIMENSION_PT) {
      const clampScale = MAX_PDF_PAGE_DIMENSION_PT / pageH;
      pageW = Math.max(1, pageW * clampScale);
      pageH = MAX_PDF_PAGE_DIMENSION_PT;
    }

    const page = pdfDoc.addPage([pageW, pageH]);
    page.drawImage(embeddedImage, {
      x: 0,
      y: 0,
      width: pageW,
      height: pageH,
    });
    return;
  }

  // 'original-px': 1px = 1pt exact pixel dimensions (clamped only to ISO 32000 14,400pt maximum)
  let pageWidth = imgW;
  let pageHeight = imgH;

  if (
    pageWidth > MAX_PDF_PAGE_DIMENSION_PT ||
    pageHeight > MAX_PDF_PAGE_DIMENSION_PT
  ) {
    const scale = MAX_PDF_PAGE_DIMENSION_PT / Math.max(pageWidth, pageHeight);
    pageWidth = Math.max(1, Math.round(pageWidth * scale));
    pageHeight = Math.max(1, Math.round(pageHeight * scale));
  }

  const page = pdfDoc.addPage([pageWidth, pageHeight]);
  page.drawImage(embeddedImage, {
    x: 0,
    y: 0,
    width: pageWidth,
    height: pageHeight,
  });
}

/**
 * Converts a single raster or vector image (`png`, `jpg`, `jfif`, `webp`, `avif`, `svg`, `tiff`, `bmp`, `ico`, `gif`)
 * into a single-page PDF via `pdf-lib` with highest visual accuracy.
 */
export async function convertImageToPdf(
  inputPath: string,
  outputPath: string,
  sourceFormat: string
): Promise<void> {
  const pdfDoc = await PDFDocument.create();
  const embeddedImage = await embedHighAccuracyImage(pdfDoc, inputPath, sourceFormat);
  addImagePageToPdf(pdfDoc, embeddedImage, 'original-px');
  const pdfBytes = await pdfDoc.save();
  await fs.writeFile(outputPath, pdfBytes);
}

export interface BatchImagePdfInput {
  inputPath: string;
  sourceFormat: string;
}

/**
 * Merges multiple images (`png`, `jpg`, `jfif`, `webp`, `avif`, `svg`, `tiff`, `bmp`, `ico`, `gif`)
 * in exact sequence into a single multi-page PDF document with lossless/high-accuracy embedding.
 */
export async function convertImagesToPdf(
  inputs: readonly BatchImagePdfInput[],
  outputPath: string,
  options?: { layout?: PdfPageLayout }
): Promise<void> {
  if (!inputs || inputs.length === 0) {
    throw new Error('No input images provided for PDF merge.');
  }

  const layout: PdfPageLayout = options?.layout ?? 'fit-image';
  const pdfDoc = await PDFDocument.create();

  for (const item of inputs) {
    const src = normalizeFormat(item.sourceFormat);
    if (!isImageFormat(src)) {
      throw new Error(`Unsupported image format in PDF merge batch: ${src.toUpperCase()}`);
    }
    const embeddedImage = await embedHighAccuracyImage(pdfDoc, item.inputPath, src);
    addImagePageToPdf(pdfDoc, embeddedImage, layout);
  }

  const pdfBytes = await pdfDoc.save();
  await fs.writeFile(outputPath, pdfBytes);
}

/**
 * Streams a plain-text (`.txt`) file line-by-line into a paginated A4 PDF document via `pdf-lib`.
 */
export async function convertTxtToPdf(inputPath: string, outputPath: string): Promise<void> {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const sanitizeLine = buildLineSanitizer(font);

  const [pageWidth, pageHeight] = PageSizes.A4;
  const maxContentWidth = pageWidth - PAGE_MARGIN_X * 2;
  const startY = pageHeight - PAGE_MARGIN_TOP - TEXT_FONT_SIZE;

  let currentPage = pdfDoc.addPage(PageSizes.A4);
  let cursorY = startY;

  const advanceNewPage = (): void => {
    currentPage = pdfDoc.addPage(PageSizes.A4);
    cursorY = startY;
  };

  const rs = createReadStream(inputPath, {
    encoding: 'utf8',
    highWaterMark: 64 * 1024,
  });
  const rl = readline.createInterface({ input: rs, crlfDelay: Infinity });

  try {
    let isFirstLine = true;

    for await (let rawLine of rl) {
      if (isFirstLine) {
        isFirstLine = false;
        if (rawLine.charCodeAt(0) === 0xfeff) {
          rawLine = rawLine.slice(1);
        }
      }

      // Support explicit form-feed (`\f`) page breaks within plain text
      const segments = rawLine.split('\f');
      for (let segIdx = 0; segIdx < segments.length; segIdx++) {
        if (segIdx > 0) {
          advanceNewPage();
        }

        const sanitized = sanitizeLine(segments[segIdx]);
        const wrappedLines = wrapTextLine(
          sanitized,
          font,
          TEXT_FONT_SIZE,
          maxContentWidth
        );

        for (const lineText of wrappedLines) {
          if (cursorY < PAGE_MARGIN_BOTTOM) {
            advanceNewPage();
          }

          if (lineText.trim().length > 0) {
            currentPage.drawText(lineText, {
              x: PAGE_MARGIN_X,
              y: cursorY,
              size: TEXT_FONT_SIZE,
              font,
              color: TEXT_COLOR,
            });
          }

          cursorY -= TEXT_LINE_HEIGHT;
        }
      }
    }
  } catch (err) {
    rl.close();
    rs.destroy();
    throw err;
  }

  const pdfBytes = await pdfDoc.save();
  await fs.writeFile(outputPath, pdfBytes);
}

/**
 * Modular Phase 2 PDF converter handler (`Image -> PDF` and `TXT -> PDF`).
 */
export async function convertPdf(
  inputPath: string,
  outputPath: string,
  sourceFormat: string,
  targetFormat: string
): Promise<void> {
  const src = normalizeFormat(sourceFormat);
  const dst = normalizeFormat(targetFormat);

  if (dst !== 'pdf' || !isPdfSourceFormat(src)) {
    throw new Error(`Unsupported PDF conversion route: ${sourceFormat} -> ${targetFormat}`);
  }

  if (src === 'txt') {
    await convertTxtToPdf(inputPath, outputPath);
    return;
  }

  if (isImageFormat(src)) {
    await convertImageToPdf(inputPath, outputPath, src);
    return;
  }

  throw new Error(`Unhandled PDF conversion source format: ${sourceFormat}`);
}
