import sharp from 'sharp';
import type { ImageFormat } from './definitions';

// Disable libvips file/memory cache to minimize RAM usage and ensure input file
// descriptors are immediately released on both Linux and Windows before fs.unlink().
sharp.cache(false);
sharp.concurrency(2);

const SUPPORTED_IMAGE_FORMATS = new Set<string>(['png', 'jpg', 'webp', 'avif']);

export function isImageFormat(fmt: string): fmt is ImageFormat {
  return SUPPORTED_IMAGE_FORMATS.has(fmt);
}

/**
 * Converts raster images (`png`, `jpg`, `webp`, `avif`) directly from `/tmp` input to `/tmp` output.
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

  // Respect EXIF orientation tags automatically
  let pipeline = sharp(inputPath, { failOn: 'error', sequentialRead: true }).rotate();

  switch (targetFormat) {
    case 'png':
      pipeline = pipeline.png({ compressionLevel: 6, adaptiveFiltering: true });
      break;
    case 'jpg':
      // Flatten alpha transparency onto white background since JPEG has no alpha channel
      pipeline = pipeline
        .flatten({ background: '#ffffff' })
        .jpeg({ quality: 85, mozjpeg: true });
      break;
    case 'webp':
      pipeline = pipeline.webp({ quality: 82, effort: 4 });
      break;
    case 'avif':
      pipeline = pipeline.avif({ quality: 65, effort: 4 });
      break;
  }

  await pipeline.toFile(outputPath);
}
