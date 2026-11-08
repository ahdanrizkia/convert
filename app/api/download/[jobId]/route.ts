import { createReadStream } from 'node:fs';
import { NextRequest, NextResponse } from 'next/server';
import {
  claimJobForDownload,
  finalizeAndPurgeJob,
  isValidUuid,
} from '@/lib/fileCleanup';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

interface RouteContext {
  params: Promise<{
    jobId: string;
  }>;
}

export async function GET(
  request: NextRequest,
  context: RouteContext
): Promise<Response> {
  const { jobId } = await context.params;

  if (!isValidUuid(jobId)) {
    return NextResponse.json(
      { error: 'Invalid download job identifier.' },
      { status: 400, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  const requestedName = request.nextUrl.searchParams.get('filename');
  const job = await claimJobForDownload(jobId, requestedName);

  if (!job) {
    return NextResponse.json(
      {
        error:
          'File is no longer on the server (already downloaded & auto-deleted, or TTL expired).',
      },
      { status: 410, headers: { 'Cache-Control': 'no-store' } }
    );
  }

  const fileStream = createReadStream(job.outputPath, {
    highWaterMark: 64 * 1024,
    autoClose: true,
  });

  const iterator = fileStream[Symbol.asyncIterator]();
  let selfDestructTriggered = false;

  /**
   * Idempotent self-destruct hook: closes the underlying file handle and immediately
   * executes `fs.unlink()` via `finalizeAndPurgeJob` on stream finish, close, error, or abort.
   */
  const selfDestruct = async (): Promise<void> => {
    if (selfDestructTriggered) return;
    selfDestructTriggered = true;

    if (!fileStream.destroyed) {
      fileStream.destroy();
    }

    if (!fileStream.closed) {
      await new Promise<void>((resolve) => {
        fileStream.once('close', () => resolve());
        setTimeout(resolve, 80);
      });
    }

    await finalizeAndPurgeJob(job.jobId, job.outputPath);
  };

  // Attach Node stream close/error listeners + request abort listener
  fileStream.once('close', () => {
    void selfDestruct();
  });
  fileStream.once('error', () => {
    void selfDestruct();
  });
  request.signal.addEventListener(
    'abort',
    () => {
      void selfDestruct();
    },
    { once: true }
  );

  const webStream = new ReadableStream<Uint8Array>({
    async pull(controller) {
      try {
        const { value, done } = await iterator.next();
        if (done) {
          // Delete `/tmp/{uuid}-output.{ext}` immediately BEFORE closing the HTTP stream
          // so the file is guaranteed purged from disk the exact moment download finishes.
          await selfDestruct();
          controller.close();
          return;
        }
        const chunk = Buffer.isBuffer(value) ? value : Buffer.from(value);
        controller.enqueue(
          new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength)
        );
      } catch (err) {
        await selfDestruct();
        controller.error(err);
      }
    },
    async cancel() {
      await selfDestruct();
    },
  });

  const asciiFallbackName = job.outputFileName.replace(/[^\x20-\x7E]/g, '_').replace(/"/g, '');
  const encodedFileName = encodeURIComponent(job.outputFileName);

  return new Response(webStream, {
    status: 200,
    headers: {
      'Content-Type': job.mimeType,
      'Content-Length': String(job.convertedSize),
      'Content-Disposition': `attachment; filename="${asciiFallbackName}"; filename*=UTF-8''${encodedFileName}`,
      'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0',
      'X-Content-Type-Options': 'nosniff',
    },
  });
}
