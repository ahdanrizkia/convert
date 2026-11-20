import crypto from 'node:crypto';
import { createWriteStream } from 'node:fs';
import fs from 'node:fs/promises';
import { Readable, Transform } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import type { ReadableStream as NodeWebReadableStream } from 'node:stream/web';
import { NextRequest, NextResponse } from 'next/server';
import {
  DEFAULT_TTL_MS,
  MAX_FILE_SIZE_BYTES,
  MAX_MERGE_FILES,
  MAX_MERGE_TOTAL_BYTES,
  MIME_TYPES,
  buildMergedPdfFileName,
  buildOutputFileName,
  isPdfMergeableImageFormat,
  normalizeFormat,
  type PdfPageLayout,
} from '@/lib/converters/definitions';
import { convertImagesToPdf, type BatchImagePdfInput } from '@/lib/converters/pdfConverter';
import { resolveConverterRoute } from '@/lib/converters/registry';
import {
  getTempBatchInputPath,
  getTempInputPath,
  getTempOutputPath,
  registerEphemeralJob,
  safeUnlink,
} from '@/lib/fileCleanup';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

class PayloadTooLargeError extends Error {
  constructor(message = 'File exceeds maximum allowed size of 50 MB.') {
    super(message);
    this.name = 'PayloadTooLargeError';
  }
}

function parsePdfPageLayout(raw: unknown): PdfPageLayout {
  if (raw === 'a4-fit' || raw === 'original-px' || raw === 'fit-image') {
    return raw;
  }
  return 'fit-image';
}

async function handleMergePdfRequest(formData: FormData): Promise<NextResponse> {
  const rawEntries = formData.getAll('files');
  const files = rawEntries.filter((entry): entry is File => entry instanceof File);

  if (files.length < 2) {
    return NextResponse.json(
      { error: 'At least 2 image files are required to merge into a PDF.' },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  if (files.length > MAX_MERGE_FILES) {
    return NextResponse.json(
      { error: `Cannot merge more than ${MAX_MERGE_FILES} images in a single batch.` },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  let declaredTotalBytes = 0;
  const validatedItems: Array<{ file: File; sourceFormat: string }> = [];

  for (const file of files) {
    if (file.size === 0) {
      return NextResponse.json(
        { error: `Uploaded file "${file.name}" is empty (0 bytes).` },
        { status: 400, headers: { 'Cache-Control': 'no-store' } }
      );
    }
    if (file.size > MAX_FILE_SIZE_BYTES) {
      return NextResponse.json(
        { error: `File "${file.name}" exceeds the 50 MB maximum upload limit.` },
        { status: 413, headers: { 'Cache-Control': 'no-store' } }
      );
    }
    declaredTotalBytes += file.size;
    if (declaredTotalBytes > MAX_MERGE_TOTAL_BYTES) {
      return NextResponse.json(
        { error: 'Total batch size exceeds the 250 MB maximum merge limit.' },
        { status: 413, headers: { 'Cache-Control': 'no-store' } }
      );
    }

    const sourceFormat = normalizeFormat(file.name);
    if (
      !sourceFormat ||
      sourceFormat === file.name.trim().toLowerCase() ||
      !isPdfMergeableImageFormat(sourceFormat)
    ) {
      return NextResponse.json(
        {
          error: `Unsupported image format for PDF merge: "${file.name}". Supported: PNG, JPG, JFIF, WEBP, AVIF, SVG, TIFF, BMP, ICO, GIF.`,
        },
        { status: 400, headers: { 'Cache-Control': 'no-store' } }
      );
    }

    validatedItems.push({ file, sourceFormat });
  }

  const layout = parsePdfPageLayout(formData.get('layout'));
  const jobId = crypto.randomUUID();
  const outputPath = getTempOutputPath(jobId, 'pdf');
  const batchInputs: BatchImagePdfInput[] = [];

  try {
    let totalBytesStreamed = 0;

    for (let idx = 0; idx < validatedItems.length; idx++) {
      const { file, sourceFormat } = validatedItems[idx];
      const inputPath = getTempBatchInputPath(jobId, idx, sourceFormat);
      batchInputs.push({ inputPath, sourceFormat });

      let fileBytesReceived = 0;
      const sizeGuardStream = new Transform({
        transform(chunk: Buffer, _encoding, callback) {
          fileBytesReceived += chunk.byteLength;
          totalBytesStreamed += chunk.byteLength;
          if (fileBytesReceived > MAX_FILE_SIZE_BYTES) {
            callback(new PayloadTooLargeError(`File "${file.name}" exceeds 50 MB limit.`));
            return;
          }
          if (totalBytesStreamed > MAX_MERGE_TOTAL_BYTES) {
            callback(
              new PayloadTooLargeError('Total batch size exceeds 250 MB maximum merge limit.')
            );
            return;
          }
          callback(null, chunk);
        },
      });

      const nodeReadable = Readable.fromWeb(
        file.stream() as unknown as NodeWebReadableStream
      );
      const tempWriteStream = createWriteStream(inputPath, { flags: 'wx' });
      await pipeline(nodeReadable, sizeGuardStream, tempWriteStream);
    }

    await convertImagesToPdf(batchInputs, outputPath, { layout });
  } catch (err: unknown) {
    await safeUnlink(outputPath);

    if (err instanceof PayloadTooLargeError) {
      return NextResponse.json(
        { error: err.message },
        { status: 413, headers: { 'Cache-Control': 'no-store' } }
      );
    }

    const message =
      err instanceof Error && err.message
        ? `PDF Merge failed: ${err.message}`
        : 'PDF Merge failed due to a malformed or corrupted image file.';

    return NextResponse.json(
      { error: message },
      { status: 422, headers: { 'Cache-Control': 'no-store' } }
    );
  } finally {
    // STRICT EPHEMERAL GUARANTEE: Immediately unlink all uploaded input files in `/tmp`
    await Promise.all(batchInputs.map((item) => safeUnlink(item.inputPath)));
  }

  let convertedSize = 0;
  try {
    const stat = await fs.stat(outputPath);
    convertedSize = stat.size;
  } catch {
    await safeUnlink(outputPath);
    return NextResponse.json(
      { error: 'Merged PDF output could not be verified in temporary storage.' },
      { status: 500, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  const outputFileName = buildMergedPdfFileName(validatedItems.map((i) => i.file.name));

  registerEphemeralJob({
    jobId,
    outputPath,
    outputFileName,
    mimeType: MIME_TYPES.pdf,
    convertedSize,
    ttlMs: DEFAULT_TTL_MS,
  });

  return NextResponse.json(
    {
      jobId,
      outputFileName,
      convertedSize,
      pageCount: validatedItems.length,
    },
    {
      status: 200,
      headers: {
        'Cache-Control': 'no-store, no-cache, must-revalidate',
      },
    }
  );
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  const isHeaderMergeMode = request.headers.get('x-convert-mode') === 'merge-pdf';

  // Fast reject on oversized Content-Length header before reading multipart stream
  const contentLengthHeader = request.headers.get('content-length');
  if (contentLengthHeader) {
    const contentLength = Number(contentLengthHeader);
    const maxAllowedBytes = isHeaderMergeMode
      ? MAX_MERGE_TOTAL_BYTES + 2 * 1024 * 1024
      : MAX_FILE_SIZE_BYTES + 512 * 1024;
    if (Number.isFinite(contentLength) && contentLength > maxAllowedBytes) {
      return NextResponse.json(
        {
          error: isHeaderMergeMode
            ? 'Total batch size exceeds the 250 MB maximum merge limit.'
            : 'File exceeds the 50 MB maximum upload limit.',
        },
        { status: 413, headers: { 'Cache-Control': 'no-store' } }
      );
    }
  }

  let formData: FormData;
  try {
    formData = await request.formData();
  } catch {
    return NextResponse.json(
      { error: 'Invalid multipart/form-data payload.' },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  if (isHeaderMergeMode || formData.get('mode') === 'merge-pdf') {
    return handleMergePdfRequest(formData);
  }

  const fileEntry = formData.get('file');
  const targetEntry = formData.get('targetFormat');

  if (!(fileEntry instanceof File)) {
    return NextResponse.json(
      { error: 'Missing uploaded file.' },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  if (fileEntry.size === 0) {
    return NextResponse.json(
      { error: 'Uploaded file is empty (0 bytes).' },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  if (fileEntry.size > MAX_FILE_SIZE_BYTES) {
    return NextResponse.json(
      { error: 'File exceeds the 50 MB maximum upload limit.' },
      { status: 413, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  if (typeof targetEntry !== 'string' || !targetEntry.trim()) {
    return NextResponse.json(
      { error: 'Missing target conversion format.' },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  const sourceFormat = normalizeFormat(fileEntry.name);
  const targetFormat = normalizeFormat(targetEntry);

  if (!sourceFormat || sourceFormat === fileEntry.name.trim().toLowerCase()) {
    return NextResponse.json(
      { error: 'Uploaded file must have a valid file extension.' },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  const route = resolveConverterRoute(sourceFormat, targetFormat);
  if (!route) {
    return NextResponse.json(
      {
        error: `Unsupported conversion route: ${sourceFormat.toUpperCase()} -> ${targetFormat.toUpperCase()}.`,
      },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  if (!route.active || !route.handler) {
    return NextResponse.json(
      {
        error: `Route ${sourceFormat.toUpperCase()} -> ${targetFormat.toUpperCase()} is registered for Phase ${route.phase} rollout.`,
      },
      { status: 501, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  const jobId = crypto.randomUUID();
  const inputPath = getTempInputPath(jobId, route.sourceFormat);
  const outputPath = getTempOutputPath(jobId, route.targetFormat);

  try {
    // Stream incoming file directly into os.tmpdir() while strictly enforcing byte cap
    let bytesReceived = 0;
    const sizeGuardStream = new Transform({
      transform(chunk: Buffer, _encoding, callback) {
        bytesReceived += chunk.byteLength;
        if (bytesReceived > MAX_FILE_SIZE_BYTES) {
          callback(new PayloadTooLargeError());
          return;
        }
        callback(null, chunk);
      },
    });

    const nodeReadable = Readable.fromWeb(
      fileEntry.stream() as unknown as NodeWebReadableStream
    );
    const tempWriteStream = createWriteStream(inputPath, { flags: 'wx' });

    await pipeline(nodeReadable, sizeGuardStream, tempWriteStream);

    // Run modular format conversion into `/tmp/{uuid}-output.{ext}`
    await route.handler(inputPath, outputPath, route.sourceFormat, route.targetFormat);
  } catch (err: unknown) {
    // Ensure partial output is cleaned up on any failure
    await safeUnlink(outputPath);

    if (err instanceof PayloadTooLargeError) {
      return NextResponse.json(
        { error: err.message },
        { status: 413, headers: { 'Cache-Control': 'no-store' } }
      );
    }

    const message =
      err instanceof Error && err.message
        ? `Conversion failed: ${err.message}`
        : 'Conversion failed due to malformed or corrupted input file.';

    return NextResponse.json(
      { error: message },
      { status: 422, headers: { 'Cache-Control': 'no-store' } }
    );
  } finally {
    // STRICT EPHEMERAL GUARANTEE: Immediately delete the original uploaded input file
    // the exact moment conversion finishes (whether succeeded or failed).
    await safeUnlink(inputPath);
  }

  // Verify output file stats and register 3-minute fallback abandonment TTL timer
  let convertedSize = 0;
  try {
    const stat = await fs.stat(outputPath);
    convertedSize = stat.size;
  } catch {
    await safeUnlink(outputPath);
    return NextResponse.json(
      { error: 'Converted output could not be verified in temporary storage.' },
      { status: 500, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  const outputFileName = buildOutputFileName(fileEntry.name, route.targetFormat);

  registerEphemeralJob({
    jobId,
    outputPath,
    outputFileName,
    mimeType: route.targetMimeType,
    convertedSize,
    ttlMs: DEFAULT_TTL_MS,
  });

  return NextResponse.json(
    {
      jobId,
      outputFileName,
      convertedSize,
    },
    {
      status: 200,
      headers: {
        'Cache-Control': 'no-store, no-cache, must-revalidate',
      },
    }
  );
}
