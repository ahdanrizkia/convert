import { spawn } from 'node:child_process';
import { constants as fsConstants } from 'node:fs';
import fs from 'node:fs/promises';
import path from 'node:path';
import ffmpegStatic from 'ffmpeg-static';
import {
  MEDIA_SOURCE_FORMATS,
  normalizeFormat,
  type MediaSourceFormat,
} from './definitions';

const SUPPORTED_MEDIA_SOURCES = new Set<string>(MEDIA_SOURCE_FORMATS);
const FFMPEG_TIMEOUT_MS = 120_000; // 2 minutes max per <=50 MB media file
const MAX_STDERR_CHARS = 16_384;

export function isMediaSourceFormat(fmt: string): fmt is MediaSourceFormat {
  return SUPPORTED_MEDIA_SOURCES.has(fmt);
}

async function isExecutableFile(candidate: string): Promise<boolean> {
  try {
    await fs.access(candidate, fsConstants.F_OK);
    const stat = await fs.stat(candidate);
    return stat.isFile() && stat.size > 0;
  } catch {
    return false;
  }
}

async function resolveFfmpegBinary(): Promise<string> {
  const candidates: string[] = [];

  if (process.env.FFMPEG_PATH?.trim()) {
    candidates.push(process.env.FFMPEG_PATH.trim());
  }

  if (typeof ffmpegStatic === 'string' && ffmpegStatic.trim().length > 0) {
    candidates.push(ffmpegStatic);
  }

  const exeName = process.platform === 'win32' ? 'ffmpeg.exe' : 'ffmpeg';
  candidates.push(
    path.join(process.cwd(), 'node_modules', 'ffmpeg-static', exeName)
  );

  for (const candidate of candidates) {
    if (await isExecutableFile(candidate)) {
      return candidate;
    }
  }

  return 'ffmpeg';
}

function formatFfmpegError(stderrText: string, exitCode: number | null): string {
  const trimmed = stderrText.trim();
  if (/Output file does not contain any stream/i.test(trimmed)) {
    return 'Input file does not contain an audio stream to extract.';
  }
  if (/Invalid data found when processing input/i.test(trimmed)) {
    return 'Malformed or corrupted media file.';
  }
  if (trimmed.length > 0) {
    const lines = trimmed
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    const lastLine = lines[lines.length - 1] ?? trimmed;
    return `FFmpeg conversion failed: ${lastLine.slice(0, 240)}`;
  }
  return `FFmpeg exited with status code ${exitCode ?? 'unknown'}.`;
}

async function runFfmpegToMp3(inputPath: string, outputPath: string): Promise<void> {
  const ffmpegBinary = await resolveFfmpegBinary();

  const args = [
    '-hide_banner',
    '-loglevel',
    'error',
    '-nostdin',
    '-y',
    '-i',
    inputPath,
    '-vn',
    '-sn',
    '-dn',
    '-map_metadata',
    '-1',
    '-c:a',
    'libmp3lame',
    '-b:a',
    '192k',
    '-f',
    'mp3',
    outputPath,
  ];

  await new Promise<void>((resolve, reject) => {
    const child = spawn(ffmpegBinary, args, {
      stdio: ['ignore', 'ignore', 'pipe'],
      windowsHide: true,
    });

    let stderrText = '';
    let settled = false;
    let timedOut = false;

    const finish = (err?: Error): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timeoutId);
      if (err) reject(err);
      else resolve();
    };

    const timeoutId = setTimeout(() => {
      timedOut = true;
      try {
        child.kill('SIGKILL');
      } catch {
        // Ignore if process already exited
      }
      // Failsafe if 'close' event does not fire promptly after SIGKILL
      const forceResolve = setTimeout(() => {
        finish(new Error('Audio conversion timed out after 120 seconds.'));
      }, 3000);
      if (typeof forceResolve.unref === 'function') {
        forceResolve.unref();
      }
    }, FFMPEG_TIMEOUT_MS);

    child.stderr?.setEncoding('utf8');
    child.stderr?.on('data', (chunk: string) => {
      if (stderrText.length < MAX_STDERR_CHARS) {
        stderrText += chunk.slice(0, MAX_STDERR_CHARS - stderrText.length);
      }
    });

    child.once('error', (err: NodeJS.ErrnoException) => {
      if (err.code === 'ENOENT') {
        finish(
          new Error(
            'FFmpeg binary not found. Ensure ffmpeg-static or system ffmpeg is available.'
          )
        );
        return;
      }
      finish(new Error(`Failed to spawn FFmpeg: ${err.message}`));
    });

    // Wait strictly for 'close' so all `/tmp` file descriptors are released by the OS
    // before `/api/convert` unlinks `inputPath` or stats `outputPath`.
    child.once('close', (code) => {
      if (timedOut) {
        finish(new Error('Audio conversion timed out after 120 seconds.'));
        return;
      }
      if (code === 0) {
        finish();
        return;
      }
      finish(new Error(formatFfmpegError(stderrText, code)));
    });
  });
}

/**
 * Modular Phase 3 media converter handler (`MP4 -> MP3` and `WAV -> MP3`).
 * Reads from `/tmp/{uuid}-input.{ext}` and writes directly to `/tmp/{uuid}-output.mp3`.
 */
export async function convertMedia(
  inputPath: string,
  outputPath: string,
  sourceFormat: string,
  targetFormat: string
): Promise<void> {
  const src = normalizeFormat(sourceFormat);
  const dst = normalizeFormat(targetFormat);

  if (dst !== 'mp3' || !isMediaSourceFormat(src)) {
    throw new Error(`Unsupported media conversion route: ${sourceFormat} -> ${targetFormat}`);
  }

  await runFfmpegToMp3(inputPath, outputPath);

  const stat = await fs.stat(outputPath);
  if (!stat.isFile() || stat.size === 0) {
    throw new Error('FFmpeg produced an empty MP3 output file.');
  }
}
