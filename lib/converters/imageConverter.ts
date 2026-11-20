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

export function resolveFfmpegBinary(): string {
  if (process.env.FFMPEG_PATH?.trim()) return process.env.FFMPEG_PATH.trim();
  if (typeof ffmpegStatic === 'string' && ffmpegStatic.trim().length > 0) return ffmpegStatic;
  const exeName = process.platform === 'win32' ? 'ffmpeg.exe' : 'ffmpeg';
  return path.join(process.cwd(), 'node_modules', 'ffmpeg-static', exeName);
}

export async function runFfmpegImageConvert(
  inputPath: string,
  outputPath: string,
  targetFormat: string
): Promise<void> {
  const ffmpegBin = resolveFfmpegBinary();
  const args = ['-hide_banner', '-loglevel', 'error', '-y', '-i', inputPath];

  switch (targetFormat) {
    case 'ico':
      // Standard icon scale with transparent padding preserving aspect ratio
      args.push(
        '-vf',
        'scale=256:256:force_original_aspect_ratio=decrease,pad=256:256:(ow-iw)/2:(oh-ih)/2:color=0x00000000'
      );
      break;
    case 'bmp':
      // Lossless 24-bit RGB Windows Bitmap
      args.push('-pix_fmt', 'bgr24');
      break;
    case 'jpg':
    case 'jfif':
      // High-quality JPEG downsampled to standard YUV420p for 10-bit/12-bit/HDR AVIF compatibility
      args.push('-vf', 'format=yuv420p', '-q:v', '2');
      break;
    case 'png':
      // Full color RGBA PNG
      args.push('-pix_fmt', 'rgba');
      break;
    case 'webp':
      args.push('-q:v', '90');
      break;
    case 'tiff':
      args.push('-pix_fmt', 'rgb24');
      break;
    case 'gif':
      // High-quality palette generation for GIF
      args.push('-vf', 'split[s0][s1];[s0]palettegen[p];[s1][p]paletteuse');
      break;
    case 'avif':
      args.push('-c:v', 'libaom-av1', '-still-picture', '1', '-crf', '25');
      break;
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

  // Sequential read is strictly supported for JPEG, PNG, WebP, and TIFF.
  // Using sequentialRead on AVIF/HEIF causes bad seek/out-of-order crashes.
  const isSequentialSafe =
    sourceFormat === 'jpg' ||
    sourceFormat === 'jfif' ||
    sourceFormat === 'png' ||
    sourceFormat === 'webp' ||
    sourceFormat === 'tiff';

  // Vector SVG inputs use 300 DPI density for ultra-crisp rasterization
  // Use failOn: 'none' to allow decoding images with harmless EXIF/container warnings
  const sharpOptions =
    sourceFormat === 'svg'
      ? { density: 300, failOn: 'none' as const }
      : { failOn: 'none' as const, ...(isSequentialSafe ? { sequentialRead: true } : {}) };

  try {
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
  } catch (sharpError) {
    // If Sharp fails (e.g. 10-bit HDR AVIF, unsupported AV1 bitstream/profile, or unusual colorspace),
    // seamlessly fall back to FFmpeg which has a complete libaom-av1 decoder.
    try {
      await runFfmpegImageConvert(inputPath, outputPath, targetFormat);
    } catch {
      // Re-throw original sharp error if fallback also fails
      throw sharpError;
    }
  }
}
