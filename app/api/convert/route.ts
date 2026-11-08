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
  buildOutputFileName,
  normalizeFormat,
} from '@/lib/converters/definitions';
import { resolveConverterRoute } from '@/lib/converters/registry';
import {
  getTempInputPath,
  getTempOutputPath,
  registerEphemeralJob,
  safeUnlink,
} from '@/lib/fileCleanup';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

class PayloadTooLargeError extends Error {
  constructor() {
    super('File exceeds maximum allowed size of 50 MB.');
    this.name = 'PayloadTooLargeError';
  }
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  // Fast reject on oversized Content-Length header before reading multipart stream
  const contentLengthHeader = request.headers.get('content-length');
  if (contentLengthHeader) {
    const contentLength = Number(contentLengthHeader);
    // Allow 512 KB overhead for multipart/form-data boundaries & fields
    if (Number.isFinite(contentLength) && contentLength > MAX_FILE_SIZE_BYTES + 512 * 1024) {
      return NextResponse.json(
        { error: 'File exceeds the 50 MB maximum upload limit.' },
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
