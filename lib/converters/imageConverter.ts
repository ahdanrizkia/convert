import { spawn } from 'node:child_process';
import path from 'node:path';
import sharp from 'sharp';
import ffmpegStatic from 'ffmpeg-static';
import type { ImageFormat } from './definitions';

// Disable libvips file/memory cache to minimize RAM usage and ensure input file
// descriptors are immediately released on both Linux and Windows before fs.unlink().
sharp.cache(false);
sharp.concurrency(2);

const SUPPORTED_IMAGE_FORMATS = new Set<string>([
  'png',
  'jpg',
  'jfif',
  'webp',
  'avif',
  'svg',
  'gif',
  'bmp',
  'ico',
  'tiff',
]);

export function isImageFormat(fmt: string): fmt is ImageFormat {
  return SUPPORTED_IMAGE_FORMATS.has(fmt);
}

function resolveFfmpegBinary(): string {
  if (process.env.FFMPEG_PATH?.trim()) return process.env.FFMPEG_PATH.trim();
  if (typeof ffmpegStatic === 'string' && ffmpegStatic.trim().length > 0) return ffmpegStatic;
  const exeName = process.platform === 'win32' ? 'ffmpeg.exe' : 'ffmpeg';
  return path.join(process.cwd(), 'node_modules', 'ffmpeg-static', exeName);
}

async function runFfmpegImageConvert(
  inputPath: string,
  outputPath: string,
  targetFormat: string
): Promise<void> {
  const ffmpegBin = resolveFfmpegBinary();
  const args = ['-hide_banner', '-loglevel', 'error', '-y', '-i', inputPath];

  if (targetFormat === 'ico') {
    // Standard icon scale with transparent padding preserving aspect ratio
    args.push(
      '-vf',
      'scale=256:256:force_original_aspect_ratio=decrease,pad=256:256:(ow-iw)/2:(oh-ih)/2:color=0x00000000'
    );
  } else if (targetFormat === 'bmp') {
    // Lossless 24-bit RGB Windows Bitmap
    args.push('-pix_fmt', 'bgr24');
  }

  args.push(outputPath);

  await new Promise<void>((resolve, reject) => {
    const child = spawn(ffmpegBin, args, { stdio: ['ignore', 'ignore', 'pipe'], windowsHide: true });
    let stderr = '';
    child.stderr?.on('data', (c: Buffer) => {
      stderr += c.toString('utf8').slice(0, 4000);
    });
    child.once('error', reject);
    child.once('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`Image transcoding failed: ${stderr.trim() || `exit code ${code}`}`));
    });
  });
}

/**
 * Converts raster & vector images directly from `/tmp` input to `/tmp` output.
 */
export async function convertImage(
  inputPath: string,
  outputPath: string,
  sourceFormat: string,
  targetFormat: string
): Promise<void> {
  if (!isImageFormat(sourceFormat) || !isImageFormat(targetFormat)) {
    throw new Error(`Unsupported image conversion route: ${sourceFormat} -> ${targetFormat}`);
  }

  // BMP and ICO formats are processed via FFmpeg for maximum format compatibility
  if (
    sourceFormat === 'bmp' ||
    sourceFormat === 'ico' ||
    targetFormat === 'bmp' ||
    targetFormat === 'ico'
  ) {
    await runFfmpegImageConvert(inputPath, outputPath, targetFormat);
    return;
  }

  // Vector SVG inputs use 300 DPI density for ultra-crisp rasterization
  const sharpOptions =
    sourceFormat === 'svg'
      ? { density: 300, failOn: 'error' as const }
      : { failOn: 'error' as const, sequentialRead: true };
  let pipeline = sharp(inputPath, sharpOptions).rotate();

  switch (targetFormat) {
    case 'png':
      pipeline = pipeline.png({ compressionLevel: 8, adaptiveFiltering: true });
      break;
    case 'jpg':
    case 'jfif':
      pipeline = pipeline
        .flatten({ background: '#ffffff' })
        .jpeg({ quality: 92, mozjpeg: true, chromaSubsampling: '4:4:4' });
      break;
    case 'webp':
      pipeline = pipeline.webp({ quality: 90, effort: 5, smartSubsample: true });
      break;
    case 'avif':
      pipeline = pipeline.avif({ quality: 75, effort: 5, chromaSubsampling: '4:4:4' });
      break;
    case 'tiff':
      pipeline = pipeline.tiff({ compression: 'deflate', quality: 90 });
      break;
    case 'gif':
      pipeline = pipeline.gif();
      break;
    default:
      pipeline = pipeline.png({ compressionLevel: 8 });
      break;
  }

  await pipeline.toFile(outputPath);
}
