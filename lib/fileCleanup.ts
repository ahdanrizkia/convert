import fs from 'node:fs/promises';
import { constants as fsConstants } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  ALLOWED_OUTPUT_EXTENSIONS,
  DEFAULT_TTL_MS,
  MIME_TYPES,
  buildOutputFileName,
} from './converters/definitions';

const UUID_REGEX =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

export interface EphemeralJobRecord {
  jobId: string;
  outputPath: string;
  outputFileName: string;
  mimeType: string;
  convertedSize: number;
  createdAt: number;
  expiresAt: number;
  timer: ReturnType<typeof setTimeout>;
}

interface GlobalEphemeralState {
  jobs: Map<string, EphemeralJobRecord>;
  inFlightDownloads: Set<string>;
}

const globalRef = globalThis as unknown as {
  __ephemeralConverterState?: GlobalEphemeralState;
};

function getState(): GlobalEphemeralState {
  if (!globalRef.__ephemeralConverterState) {
    globalRef.__ephemeralConverterState = {
      jobs: new Map<string, EphemeralJobRecord>(),
      inFlightDownloads: new Set<string>(),
    };
  }
  return globalRef.__ephemeralConverterState;
}

/**
 * Validates that a jobId is strictly a canonical cryptographic UUID.
 * Prevents any path traversal when resolving files in os.tmpdir().
 */
export function isValidUuid(value: string): boolean {
  return UUID_REGEX.test(value);
}

/**
 * Resolves a temporary input path in os.tmpdir(): `/tmp/{uuid}-input.{ext}`
 */
export function getTempInputPath(jobId: string, sourceFormat: string): string {
  if (!isValidUuid(jobId)) {
    throw new Error('Invalid cryptographic job UUID.');
  }
  const safeExt = sourceFormat.replace(/[^a-z0-9]/gi, '').toLowerCase();
  return path.join(os.tmpdir(), `${jobId}-input.${safeExt}`);
}

/**
 * Resolves a temporary output path in os.tmpdir(): `/tmp/{uuid}-output.{ext}`
 */
export function getTempOutputPath(jobId: string, targetFormat: string): string {
  if (!isValidUuid(jobId)) {
    throw new Error('Invalid cryptographic job UUID.');
  }
  const safeExt = targetFormat.replace(/[^a-z0-9]/gi, '').toLowerCase();
  return path.join(os.tmpdir(), `${jobId}-output.${safeExt}`);
}

/**
 * Checks whether a file currently exists in the filesystem.
 */
export async function fileExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath, fsConstants.F_OK);
    return true;
  } catch {
    return false;
  }
}

/**
 * Safely and idempotently deletes a temporary file inside `os.tmpdir()`.
 * Includes brief retry handling for Windows file descriptor release (`EBUSY` / `EPERM`)
 * and ignores `ENOENT` if the file was already purged.
 */
export async function safeUnlink(filePath: string): Promise<boolean> {
  if (!filePath) return false;

  const resolvedTarget = path.resolve(filePath);
  const resolvedTmpRoot = path.resolve(os.tmpdir());

  // Defense-in-depth: never unlink anything outside os.tmpdir()
  if (
    !resolvedTarget.startsWith(resolvedTmpRoot + path.sep) &&
    resolvedTarget !== resolvedTmpRoot
  ) {
    throw new Error('Refusing to unlink path outside os.tmpdir().');
  }

  const maxAttempts = 4;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      await fs.unlink(resolvedTarget);
      return true;
    } catch (err: unknown) {
      const code = (err as NodeJS.ErrnoException)?.code;
      if (code === 'ENOENT') {
        return false; // Already deleted cleanly
      }
      if ((code === 'EBUSY' || code === 'EPERM') && attempt < maxAttempts) {
        await new Promise((resolve) => setTimeout(resolve, 40 * attempt));
        continue;
      }
      return false;
    }
  }
  return false;
}

/**
 * Registers a newly converted output file with a strict fallback abandonment TTL timer.
 * If the user closes their tab or loses Wi-Fi before downloading, the timer automatically
 * checks if the output file still exists and purges it via `safeUnlink`.
 */
export function registerEphemeralJob(params: {
  jobId: string;
  outputPath: string;
  outputFileName: string;
  mimeType: string;
  convertedSize: number;
  ttlMs?: number;
}): EphemeralJobRecord {
  const {
    jobId,
    outputPath,
    outputFileName,
    mimeType,
    convertedSize,
    ttlMs = DEFAULT_TTL_MS,
  } = params;

  const state = getState();
  const existing = state.jobs.get(jobId);
  if (existing) {
    clearTimeout(existing.timer);
  }

  const now = Date.now();
  const timer = setTimeout(async () => {
    const currentState = getState();
    currentState.jobs.delete(jobId);
    // If a download stream is actively reading it right as TTL hits, let the stream finish hook delete it
    if (currentState.inFlightDownloads.has(jobId)) {
      return;
    }
    if (await fileExists(outputPath)) {
      await safeUnlink(outputPath);
    }
  }, ttlMs);

  // Allow clean process exit in CLI / test environments while staying active in server runtime
  if (typeof timer.unref === 'function') {
    timer.unref();
  }

  const record: EphemeralJobRecord = {
    jobId,
    outputPath,
    outputFileName,
    mimeType,
    convertedSize,
    createdAt: now,
    expiresAt: now + ttlMs,
    timer,
  };

  state.jobs.set(jobId, record);
  return record;
}

export interface ClaimedDownloadJob {
  jobId: string;
  outputPath: string;
  outputFileName: string;
  mimeType: string;
  convertedSize: number;
}

/**
 * Atomically claims a converted `/tmp` file for streaming download.
 * Cancels the fallback TTL timer because the download response lifecycle now guarantees
 * immediate deletion upon stream completion or client abort.
 */
export async function claimJobForDownload(
  jobId: string,
  requestedFileName?: string | null
): Promise<ClaimedDownloadJob | null> {
  if (!isValidUuid(jobId)) {
    return null;
  }

  const state = getState();
  if (state.inFlightDownloads.has(jobId)) {
    return null; // Already being downloaded and scheduled for immediate destruction
  }

  const inMemory = state.jobs.get(jobId);
  if (inMemory) {
    clearTimeout(inMemory.timer);
    state.jobs.delete(jobId);

    try {
      const stat = await fs.stat(inMemory.outputPath);
      if (!stat.isFile()) {
        await safeUnlink(inMemory.outputPath);
        return null;
      }
      state.inFlightDownloads.add(jobId);
      return {
        jobId,
        outputPath: inMemory.outputPath,
        outputFileName: inMemory.outputFileName,
        mimeType: inMemory.mimeType,
        convertedSize: stat.size,
      };
    } catch {
      return null;
    }
  }

  // Stateless filesystem fallback (e.g., if dev HMR reloaded module state before TTL fired)
  for (const ext of ALLOWED_OUTPUT_EXTENSIONS) {
    const candidatePath = getTempOutputPath(jobId, ext);
    try {
      const stat = await fs.stat(candidatePath);
      if (stat.isFile()) {
        state.inFlightDownloads.add(jobId);
        const safeName = requestedFileName
          ? buildOutputFileName(requestedFileName, ext)
          : `converted-${jobId.slice(0, 8)}.${ext}`;
        return {
          jobId,
          outputPath: candidatePath,
          outputFileName: safeName,
          mimeType: MIME_TYPES[ext] ?? 'application/octet-stream',
          convertedSize: stat.size,
        };
      }
    } catch {
      // Continue checking next allowed extension
    }
  }

  return null;
}

/**
 * Immediately purges a claimed job's output file from `/tmp` and clears all tracking state.
 */
export async function finalizeAndPurgeJob(
  jobId: string,
  outputPath: string
): Promise<void> {
  const state = getState();
  const existing = state.jobs.get(jobId);
  if (existing) {
    clearTimeout(existing.timer);
    state.jobs.delete(jobId);
  }
  try {
    await safeUnlink(outputPath);
  } finally {
    state.inFlightDownloads.delete(jobId);
  }
}
