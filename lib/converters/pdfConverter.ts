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
import { normalizeFormat, type PdfSourceFormat, PDF_SOURCE_FORMATS } from './definitions';

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
 * Converts raster images (`png`, `jpg`, `jfif`, `webp`, `avif`) into a single-page PDF via `pdf-lib`.
 * Uses `sharp` to normalize EXIF orientation and transcode WebP/AVIF buffers for `pdf-lib`.
 */
export async function convertImageToPdf(
  inputPath: string,
  outputPath: string,
  sourceFormat: string
): Promise<void> {
  // Read into buffer so the `/tmp` input file descriptor closes immediately
  let inputBuffer = await fs.readFile(inputPath);

  // If BMP or ICO, transcode buffer to PNG via FFmpeg first
  if (sourceFormat === 'bmp' || sourceFormat === 'ico') {
    const { spawn } = await import('node:child_process');
    const ffmpegStatic = (await import('ffmpeg-static')).default;
    const exeName = process.platform === 'win32' ? 'ffmpeg.exe' : 'ffmpeg';
    const ffmpegBin =
      process.env.FFMPEG_PATH?.trim() ||
      (typeof ffmpegStatic === 'string' ? ffmpegStatic : path.join(process.cwd(), 'node_modules', 'ffmpeg-static', exeName));

    const pngChunks: Buffer[] = [];
    await new Promise<void>((resolve, reject) => {
      const child = spawn(ffmpegBin, ['-hide_banner', '-loglevel', 'error', '-y', '-i', inputPath, '-f', 'image2pipe', '-vcodec', 'png', '-'], {
        stdio: ['ignore', 'pipe', 'pipe'],
        windowsHide: true,
      });
      child.stdout.on('data', (d: Buffer) => pngChunks.push(d));
      child.once('error', reject);
      child.once('close', (code) => {
        if (code === 0) resolve();
        else reject(new Error(`Failed to decode ${sourceFormat} for PDF`));
      });
    });
    inputBuffer = Buffer.concat(pngChunks);
    sourceFormat = 'png';
  }

  const sharpOptions =
    sourceFormat === 'svg'
      ? { density: 300, failOn: 'error' as const }
      : { failOn: 'error' as const };
  const metadata = await sharp(inputBuffer, sharpOptions).metadata();

  const needsExifRotation =
    typeof metadata.orientation === 'number' && metadata.orientation > 1;
  const isCmyk = metadata.space === 'cmyk';
  const isJpeg = sourceFormat === 'jpg' || sourceFormat === 'jfif';

  const pdfDoc = await PDFDocument.create();
  let embeddedImage: PDFImage;

  if (isJpeg && !needsExifRotation && !isCmyk) {
    try {
      embeddedImage = await pdfDoc.embedJpg(inputBuffer);
    } catch {
      const normalizedJpg = await sharp(inputBuffer, sharpOptions)
        .rotate()
        .flatten({ background: '#ffffff' })
        .jpeg({ quality: 94, mozjpeg: true, chromaSubsampling: '4:4:4' })
        .toBuffer();
      embeddedImage = await pdfDoc.embedJpg(normalizedJpg);
    }
  } else if (sourceFormat === 'png' && !needsExifRotation) {
    try {
      embeddedImage = await pdfDoc.embedPng(inputBuffer);
    } catch {
      const normalizedPng = await sharp(inputBuffer, sharpOptions)
        .rotate()
        .png({ compressionLevel: 8 })
        .toBuffer();
      embeddedImage = await pdfDoc.embedPng(normalizedPng);
    }
  } else {
    const pipeline = sharp(inputBuffer, sharpOptions).rotate();
    if (metadata.hasAlpha && !isJpeg) {
      const pngBuffer = await pipeline.png({ compressionLevel: 8 }).toBuffer();
      embeddedImage = await pdfDoc.embedPng(pngBuffer);
    } else {
      const jpgBuffer = await pipeline
        .flatten({ background: '#ffffff' })
        .jpeg({ quality: 94, mozjpeg: true, chromaSubsampling: '4:4:4' })
        .toBuffer();
      embeddedImage = await pdfDoc.embedJpg(jpgBuffer);
    }
  }

  let pageWidth = embeddedImage.width;
  let pageHeight = embeddedImage.height;

  if (pageWidth <= 0 || pageHeight <= 0) {
    throw new Error('Invalid image dimensions for PDF generation.');
  }

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
