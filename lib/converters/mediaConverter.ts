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
  candidates.push(path.join(process.cwd(), 'node_modules', 'ffmpeg-static', exeName));

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
    return 'Input file does not contain an audio/video stream to extract.';
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

async function runFfmpegMediaConvert(
  inputPath: string,
  outputPath: string,
  sourceFormat: string,
  targetFormat: string
): Promise<void> {
  const ffmpegBinary = await resolveFfmpegBinary();
  let args: string[] = [];

  if (targetFormat === 'mp3') {
    args = [
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
      '320k',
      '-qscale:a',
      '0',
      '-f',
      'mp3',
      outputPath,
    ];
  } else if (targetFormat === 'wav') {
    args = [
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
      'pcm_s16le',
      '-f',
      'wav',
      outputPath,
    ];
  } else if (targetFormat === 'mp4') {
    if (sourceFormat === 'gif') {
      args = [
        '-hide_banner',
        '-loglevel',
        'error',
        '-nostdin',
        '-y',
        '-i',
        inputPath,
        '-c:v',
        'libx264',
        '-crf',
        '18',
        '-preset',
        'medium',
        '-movflags',
        '+faststart',
        '-pix_fmt',
        'yuv420p',
        '-vf',
        'scale=trunc(iw/2)*2:trunc(ih/2)*2',
        outputPath,
      ];
    } else if (
      sourceFormat === 'mp3' ||
      sourceFormat === 'wav' ||
      sourceFormat === 'ogg' ||
      sourceFormat === 'aac' ||
      sourceFormat === 'm4a'
    ) {
      args = [
        '-hide_banner',
        '-loglevel',
        'error',
        '-nostdin',
        '-y',
        '-f',
        'lavfi',
        '-i',
        'color=c=black:s=1280x720:r=1',
        '-i',
        inputPath,
        '-c:v',
        'libx264',
        '-tune',
        'stillimage',
        '-c:a',
        'aac',
        '-b:a',
        '320k',
        '-pix_fmt',
        'yuv420p',
        '-shortest',
        outputPath,
      ];
    } else {
      args = [
        '-hide_banner',
        '-loglevel',
        'error',
        '-nostdin',
        '-y',
        '-i',
        inputPath,
        '-c:v',
        'libx264',
        '-crf',
        '18',
        '-preset',
        'medium',
        '-vf',
        'scale=trunc(iw/2)*2:trunc(ih/2)*2',
        '-pix_fmt',
        'yuv420p',
        '-c:a',
        'aac',
        '-b:a',
        '256k',
        '-movflags',
        '+faststart',
        outputPath,
      ];
    }
  } else {
    throw new Error(`Unsupported media target format: ${targetFormat}`);
  }

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
        // Ignore if already dead
      }
      const forceResolve = setTimeout(() => {
        finish(new Error('Media conversion timed out after 120 seconds.'));
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

    child.once('close', (code) => {
      if (timedOut) {
        finish(new Error('Media conversion timed out after 120 seconds.'));
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
 * Universal media converter handler (`mp4`, `mp3`, `wav`, `ogg`, `aac`, `m4a`, `mov`, `webm`, `mkv`, `gif`).
 */
export async function convertMedia(
  inputPath: string,
  outputPath: string,
  sourceFormat: string,
  targetFormat: string
): Promise<void> {
  const src = normalizeFormat(sourceFormat);
  const dst = normalizeFormat(targetFormat);

  if (!isMediaSourceFormat(src)) {
    throw new Error(`Unsupported media source format: ${sourceFormat}`);
  }

  await runFfmpegMediaConvert(inputPath, outputPath, src, dst);

  const stat = await fs.stat(outputPath);
  if (!stat.isFile() || stat.size === 0) {
    throw new Error(`FFmpeg produced an empty output file for ${dst.toUpperCase()}.`);
  }
}
