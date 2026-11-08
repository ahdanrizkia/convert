/** @type {import('next').NextConfig} */
const nextConfig = {
  serverExternalPackages: [
    'sharp',
    'ffmpeg-static',
    '@napi-rs/canvas',
    'pdfjs-dist',
    'pdf-parse',
    'mammoth',
    'xlsx',
    'docx',
    'pptxgenjs',
    'jszip',
  ],
  poweredByHeader: false,
  reactStrictMode: true,
};

export default nextConfig;
