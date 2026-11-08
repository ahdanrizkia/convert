import fs from 'node:fs/promises';
import sharp from 'sharp';
import { normalizeFormat } from './definitions';
import { convertImageToPdf, convertTxtToPdf } from './pdfConverter';

// Keep libvips memory flat
sharp.cache(false);

/**
 * Text sanitizer for standard WinAnsi font in pdf-lib (Helvetica)
 */
function sanitizeWinAnsiText(str: string): string {
  return str
    .replace(/\t/g, '    ')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/\u2014/g, '--')
    .replace(/\u2013/g, '-')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[^\x20-\x7E\xA0-\xFF\n\r]/g, '?');
}

/**
 * Helper to split text into wrapped lines fitting within a given width
 */
function wrapLine(
  line: string,
  widthOfText: (text: string) => number,
  maxWidth: number
): string[] {
  if (line.length === 0) return [''];
  if (widthOfText(line) <= maxWidth) return [line];

  const words = line.match(/\S+|\s+/g) ?? [line];
  const out: string[] = [];
  let cur = '';

  for (const word of words) {
    if (!cur) {
      if (widthOfText(word) <= maxWidth) {
        cur = word;
      } else {
        // Break long unbreakable token
        let chunk = '';
        for (const ch of word) {
          if (widthOfText(chunk + ch) > maxWidth && chunk) {
            out.push(chunk);
            chunk = ch;
          } else {
            chunk += ch;
          }
        }
        cur = chunk;
      }
      continue;
    }

    if (widthOfText(cur + word) <= maxWidth) {
      cur += word;
    } else {
      out.push(cur.trimEnd());
      cur = /^\s+$/.test(word) ? '' : word;
    }
  }

  if (cur.trim().length > 0) {
    out.push(cur.trimEnd());
  }

  return out.length > 0 ? out : [''];
}

/* ==========================================================================
   1. Images to Office/PDF (JPG, PNG, WEBP, AVIF -> PDF, DOCX, PPTX)
   ========================================================================== */

async function convertImageToDocx(
  inputPath: string,
  outputPath: string,
  sourceFormat: string
): Promise<void> {
  const { Document, Packer, Paragraph, ImageRun } = await import('docx');

  const rawBytes = await fs.readFile(inputPath);
  let imgBuffer: Uint8Array = new Uint8Array(rawBytes);
  let ext = sourceFormat;

  // Transcode webp/avif to png for Word OpenXML compatibility
  if (ext === 'webp' || ext === 'avif') {
    imgBuffer = new Uint8Array(await sharp(imgBuffer).png({ compressionLevel: 6 }).toBuffer());
    ext = 'png';
  }

  const meta = await sharp(imgBuffer).metadata();
  const rawW = meta.width ?? 600;
  const rawH = meta.height ?? 400;

  // Fit image to max 540pt width on portrait page
  const maxW = 540;
  const scale = Math.min(1, maxW / rawW);
  const width = Math.max(1, Math.round(rawW * scale));
  const height = Math.max(1, Math.round(rawH * scale));

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            children: [
              new ImageRun({
                data: imgBuffer,
                transformation: { width, height },
                type: ext === 'jpg' ? 'jpg' : 'png',
              }),
            ],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(outputPath, new Uint8Array(buffer));
}

async function convertImageToPptx(
  inputPath: string,
  outputPath: string,
  sourceFormat: string
): Promise<void> {
  const pptxgen = (await import('pptxgenjs')).default;

  const rawBytes = await fs.readFile(inputPath);
  let imgBuffer: Uint8Array = new Uint8Array(rawBytes);
  let mimeType = sourceFormat === 'jpg' ? 'image/jpeg' : 'image/png';

  if (sourceFormat === 'webp' || sourceFormat === 'avif') {
    imgBuffer = new Uint8Array(await sharp(imgBuffer).png({ compressionLevel: 6 }).toBuffer());
    mimeType = 'image/png';
  }

  const meta = await sharp(imgBuffer).metadata();
  const rawW = meta.width || 800;
  const rawH = meta.height || 600;
  const imgRatio = rawW / rawH;

  // 16:9 Slide is 10.0 x 5.625 inches
  let w = 9.0;
  let h = w / imgRatio;
  if (h > 5.0) {
    h = 5.0;
    w = h * imgRatio;
  }
  const x = (10.0 - w) / 2;
  const y = (5.625 - h) / 2;

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  const slide = pres.addSlide();

  slide.addImage({
    data: `data:${mimeType};base64,${Buffer.from(imgBuffer).toString('base64')}`,
    x: Number(x.toFixed(2)),
    y: Number(y.toFixed(2)),
    w: Number(w.toFixed(2)),
    h: Number(h.toFixed(2)),
  });

  const buffer = (await pres.write({ outputType: 'nodebuffer' })) as Buffer;
  await fs.writeFile(outputPath, new Uint8Array(buffer));
}

/* ==========================================================================
   2. Word (DOCX) Conversions (DOCX -> PDF, TXT, HTML)
   ========================================================================== */

async function convertDocxToPdf(inputPath: string, outputPath: string): Promise<void> {
  const mammoth = await import('mammoth');
  const { PDFDocument, PageSizes, StandardFonts, rgb } = await import('pdf-lib');

  const { value } = await mammoth.extractRawText({ path: inputPath });
  const rawLines = (value || '').split(/\r?\n/);

  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const fontSize = 10.5;
  const lineHeight = 15;
  const marginX = 50;
  const marginTop = 50;
  const marginBottom = 50;

  const [pageW, pageH] = PageSizes.A4;
  const maxW = pageW - marginX * 2;
  const startY = pageH - marginTop - fontSize;

  let page = doc.addPage(PageSizes.A4);
  let cursorY = startY;

  const widthOfText = (t: string) => font.widthOfTextAtSize(t, fontSize);

  for (const rawLine of rawLines) {
    const clean = sanitizeWinAnsiText(rawLine);
    const wrapped = wrapLine(clean, widthOfText, maxW);

    for (const line of wrapped) {
      if (cursorY < marginBottom) {
        page = doc.addPage(PageSizes.A4);
        cursorY = startY;
      }

      if (line.trim().length > 0) {
        page.drawText(line, {
          x: marginX,
          y: cursorY,
          size: fontSize,
          font,
          color: rgb(0.1, 0.1, 0.1),
        });
      }

      cursorY -= lineHeight;
    }
  }

  const pdfBytes = await doc.save();
  await fs.writeFile(outputPath, pdfBytes);
}

async function convertDocxToTxt(inputPath: string, outputPath: string): Promise<void> {
  const mammoth = await import('mammoth');
  const { value } = await mammoth.extractRawText({ path: inputPath });
  await fs.writeFile(outputPath, value || '', 'utf8');
}

async function convertDocxToHtml(inputPath: string, outputPath: string): Promise<void> {
  const mammoth = await import('mammoth');
  const { value } = await mammoth.convertToHtml({ path: inputPath });

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Converted Document</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      max-width: 800px;
      margin: 2rem auto;
      padding: 1.5rem;
      line-height: 1.6;
      color: #1a1a1a;
      background-color: #fafafa;
    }
    p { margin-bottom: 1em; }
    table { border-collapse: collapse; width: 100%; margin: 1em 0; }
    td, th { border: 1px solid #ccc; padding: 6px 10px; }
  </style>
</head>
<body>
${value || '<p>(Empty Document)</p>'}
</body>
</html>`;

  await fs.writeFile(outputPath, html, 'utf8');
}

/* ==========================================================================
   3. PDF Conversions (PDF -> DOCX, PNG, JPG, TXT)
   ========================================================================== */

async function convertPdfToDocx(inputPath: string, outputPath: string): Promise<void> {
  const { PDFParse } = await import('pdf-parse');
  const { Document, Packer, Paragraph, TextRun, HeadingLevel } = await import('docx');

  const pdfBuffer = await fs.readFile(inputPath);
  const parser = new PDFParse({ data: pdfBuffer });
  let parseResult: { pages?: Array<{ text?: string; num?: number }>; text?: string } | null =
    null;

  try {
    parseResult = await parser.getText();
  } finally {
    await parser.destroy();
    delete (globalThis as unknown as { pdfjsLib?: unknown }).pdfjsLib;
  }

  const sectionsChildren: InstanceType<typeof Paragraph>[] = [];
  const pages = parseResult?.pages ?? [];

  if (pages.length > 0) {
    for (const pg of pages) {
      if (pages.length > 1) {
        sectionsChildren.push(
          new Paragraph({
            text: `Page ${pg.num ?? 1}`,
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 240, after: 120 },
          })
        );
      }
      const lines = (pg.text || '').split(/\r?\n/);
      for (const line of lines) {
        sectionsChildren.push(
          new Paragraph({
            children: [new TextRun(line)],
            spacing: { after: 80 },
          })
        );
      }
    }
  } else {
    const rawLines = (parseResult?.text || '').split(/\r?\n/);
    for (const line of rawLines) {
      sectionsChildren.push(
        new Paragraph({
          children: [new TextRun(line)],
          spacing: { after: 80 },
        })
      );
    }
  }

  if (sectionsChildren.length === 0) {
    sectionsChildren.push(new Paragraph('(Empty PDF)'));
  }

  const doc = new Document({
    sections: [{ children: sectionsChildren }],
  });

  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(outputPath, new Uint8Array(buffer));
}

async function convertPdfToImage(
  inputPath: string,
  outputPath: string,
  targetFormat: 'png' | 'jpg'
): Promise<void> {
  const { PDFParse } = await import('pdf-parse');
  const pdfBuffer = await fs.readFile(inputPath);
  const parser = new PDFParse({ data: new Uint8Array(pdfBuffer) });

  let pngBytes: Uint8Array | null = null;
  try {
    const screenshot = await parser.getScreenshot({ pageNumber: 1, scale: 2.0 });
    if (screenshot?.pages?.[0]?.data) {
      pngBytes = screenshot.pages[0].data;
    }
  } finally {
    await parser.destroy();
  }

  if (!pngBytes) {
    throw new Error('Failed to render PDF page 1 to image.');
  }

  if (targetFormat === 'jpg') {
    const jpgBytes = await sharp(pngBytes)
      .flatten({ background: '#ffffff' })
      .jpeg({ quality: 90, mozjpeg: true })
      .toBuffer();
    await fs.writeFile(outputPath, new Uint8Array(jpgBytes));
  } else {
    await fs.writeFile(outputPath, new Uint8Array(pngBytes));
  }
}

async function convertPdfToTxt(inputPath: string, outputPath: string): Promise<void> {
  const { PDFParse } = await import('pdf-parse');
  const pdfBuffer = await fs.readFile(inputPath);
  const parser = new PDFParse({ data: pdfBuffer });
  let parseResult: { text?: string } | null = null;

  try {
    parseResult = await parser.getText();
  } finally {
    await parser.destroy();
    delete (globalThis as unknown as { pdfjsLib?: unknown }).pdfjsLib;
  }

  await fs.writeFile(outputPath, parseResult?.text || '', 'utf8');
}

/* ==========================================================================
   4. Spreadsheet (XLSX / CSV) Conversions (XLSX -> PDF, CSV, JSON; CSV/JSON -> XLSX)
   ========================================================================== */

async function convertXlsxToPdf(inputPath: string, outputPath: string): Promise<void> {
  const XLSX = await import('xlsx');
  const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib');

  const wb = XLSX.readFile(inputPath);
  const sheetName = wb.SheetNames[0] || 'Sheet1';
  const sheet = wb.Sheets[sheetName];
  const rawRows: unknown[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const fontBold = await doc.embedFont(StandardFonts.HelveticaBold);

  // Landscape A4 dimensions in points: 841.89 x 595.28
  const pageW = 841.89;
  const pageH = 595.28;
  const marginX = 40;
  const marginTop = 40;
  const marginBottom = 40;
  const contentW = pageW - marginX * 2;
  const rowHeight = 22;
  const fontSize = 8.5;

  let maxCols = 0;
  for (const row of rawRows) {
    if (Array.isArray(row) && row.length > maxCols) {
      maxCols = row.length;
    }
  }
  maxCols = Math.min(Math.max(maxCols, 1), 12); // cap table width to 12 cols for clean layout
  const colWidth = contentW / maxCols;

  let page = doc.addPage([pageW, pageH]);
  let cursorY = pageH - marginTop;

  // Draw title banner
  page.drawText(`Sheet: ${sanitizeWinAnsiText(sheetName)}`, {
    x: marginX,
    y: cursorY - 12,
    size: 11,
    font: fontBold,
    color: rgb(0.1, 0.1, 0.1),
  });
  cursorY -= 28;

  for (let rIdx = 0; rIdx < rawRows.length; rIdx++) {
    if (cursorY - rowHeight < marginBottom) {
      page = doc.addPage([pageW, pageH]);
      cursorY = pageH - marginTop;
    }

    const row = Array.isArray(rawRows[rIdx]) ? (rawRows[rIdx] as unknown[]) : [];
    const isHeader = rIdx === 0;

    // Row background
    if (isHeader) {
      page.drawRectangle({
        x: marginX,
        y: cursorY - rowHeight,
        width: contentW,
        height: rowHeight,
        color: rgb(0.92, 0.92, 0.92),
      });
    } else if (rIdx % 2 === 1) {
      page.drawRectangle({
        x: marginX,
        y: cursorY - rowHeight,
        width: contentW,
        height: rowHeight,
        color: rgb(0.98, 0.98, 0.98),
      });
    }

    // Cell borders & values
    for (let cIdx = 0; cIdx < maxCols; cIdx++) {
      const cellX = marginX + cIdx * colWidth;
      const cellY = cursorY - rowHeight;

      page.drawRectangle({
        x: cellX,
        y: cellY,
        width: colWidth,
        height: rowHeight,
        borderWidth: 0.5,
        borderColor: rgb(0.8, 0.8, 0.8),
      });

      const cellVal = row[cIdx] !== undefined && row[cIdx] !== null ? String(row[cIdx]) : '';
      const cleanVal = sanitizeWinAnsiText(cellVal);

      // Truncate to fit cell
      let displayVal = cleanVal;
      const maxTextW = colWidth - 8;
      const currentFont = isHeader ? fontBold : font;
      while (displayVal.length > 0 && currentFont.widthOfTextAtSize(displayVal, fontSize) > maxTextW) {
        displayVal = displayVal.slice(0, -1);
      }
      if (displayVal.length < cleanVal.length && displayVal.length > 3) {
        displayVal = displayVal.slice(0, -2) + '..';
      }

      if (displayVal) {
        page.drawText(displayVal, {
          x: cellX + 4,
          y: cellY + 6,
          size: fontSize,
          font: currentFont,
          color: rgb(0.1, 0.1, 0.1),
        });
      }
    }

    cursorY -= rowHeight;
  }

  const pdfBytes = await doc.save();
  await fs.writeFile(outputPath, pdfBytes);
}

async function convertXlsxToCsv(inputPath: string, outputPath: string): Promise<void> {
  const XLSX = await import('xlsx');
  const wb = XLSX.readFile(inputPath);
  const sheetName = wb.SheetNames[0];
  const csv = sheetName ? XLSX.utils.sheet_to_csv(wb.Sheets[sheetName]) : '';
  await fs.writeFile(outputPath, csv || '', 'utf8');
}

async function convertXlsxToJson(inputPath: string, outputPath: string): Promise<void> {
  const XLSX = await import('xlsx');
  const wb = XLSX.readFile(inputPath);
  const sheetName = wb.SheetNames[0];
  const json = sheetName ? XLSX.utils.sheet_to_json(wb.Sheets[sheetName]) : [];
  await fs.writeFile(outputPath, JSON.stringify(json, null, 2), 'utf8');
}

async function convertCsvToXlsx(inputPath: string, outputPath: string): Promise<void> {
  const XLSX = await import('xlsx');
  const csvContent = await fs.readFile(inputPath, 'utf8');
  const wb = XLSX.read(csvContent, { type: 'string' });
  XLSX.writeFile(wb, outputPath);
}

async function convertJsonToXlsx(inputPath: string, outputPath: string): Promise<void> {
  const XLSX = await import('xlsx');
  const jsonContent = await fs.readFile(inputPath, 'utf8');
  const parsed = JSON.parse(jsonContent);
  const dataArray = Array.isArray(parsed) ? parsed : [parsed];
  const ws = XLSX.utils.json_to_sheet(dataArray);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, outputPath);
}

/* ==========================================================================
   5. Presentation (PPTX) Conversions (PPTX -> PDF, TXT)
   ========================================================================== */

interface ParsedSlide {
  num: number;
  texts: string[];
}

async function extractPptxSlides(inputPath: string): Promise<ParsedSlide[]> {
  const JSZip = (await import('jszip')).default;
  const zipBuffer = await fs.readFile(inputPath);
  const zip = await JSZip.loadAsync(zipBuffer);

  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)?.[0] || '0', 10);
      const numB = parseInt(b.match(/\d+/)?.[0] || '0', 10);
      return numA - numB;
    });

  const slides: ParsedSlide[] = [];

  for (let idx = 0; idx < slideFiles.length; idx++) {
    const filename = slideFiles[idx];
    const xml = await zip.files[filename].async('text');
    const regex = /<a:t[^>]*>(.*?)<\/a:t>/gs;
    const texts: string[] = [];
    let m: RegExpExecArray | null;

    while ((m = regex.exec(xml)) !== null) {
      const unescaped = m[1]
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .trim();
      if (unescaped) {
        texts.push(unescaped);
      }
    }

    slides.push({
      num: idx + 1,
      texts,
    });
  }

  return slides;
}

async function convertPptxToTxt(inputPath: string, outputPath: string): Promise<void> {
  const slides = await extractPptxSlides(inputPath);
  const blocks: string[] = [];

  for (const slide of slides) {
    blocks.push(`=== SLIDE ${slide.num} ===\n${slide.texts.join('\n')}`);
  }

  await fs.writeFile(outputPath, blocks.join('\n\n') || '(Empty Presentation)', 'utf8');
}

async function convertPptxToPdf(inputPath: string, outputPath: string): Promise<void> {
  const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib');
  const slides = await extractPptxSlides(inputPath);

  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const fontBold = await doc.embedFont(StandardFonts.HelveticaBold);

  // 16:9 Landscape Slide dimensions in points: 960 x 540
  const slideW = 960;
  const slideH = 540;
  const marginX = 60;
  const marginTop = 50;
  const marginBottom = 50;
  const maxW = slideW - marginX * 2;

  const widthOfText = (t: string) => font.widthOfTextAtSize(t, 13);

  if (slides.length === 0) {
    const page = doc.addPage([slideW, slideH]);
    page.drawText('(Empty Presentation)', {
      x: marginX,
      y: slideH / 2,
      size: 16,
      font,
      color: rgb(0.4, 0.4, 0.4),
    });
  } else {
    for (const slide of slides) {
      const page = doc.addPage([slideW, slideH]);

      // Border frame around slide
      page.drawRectangle({
        x: 30,
        y: 30,
        width: slideW - 60,
        height: slideH - 60,
        borderWidth: 1,
        borderColor: rgb(0.85, 0.85, 0.85),
      });

      // Slide Title Banner
      page.drawText(`Slide ${slide.num}`, {
        x: marginX,
        y: slideH - marginTop - 16,
        size: 18,
        font: fontBold,
        color: rgb(0.1, 0.1, 0.1),
      });

      let cursorY = slideH - marginTop - 46;

      for (const text of slide.texts) {
        const clean = sanitizeWinAnsiText(text);
        const wrapped = wrapLine(clean, widthOfText, maxW);

        for (const line of wrapped) {
          if (cursorY < marginBottom + 20) break;
          page.drawText(line, {
            x: marginX,
            y: cursorY,
            size: 13,
            font,
            color: rgb(0.2, 0.2, 0.2),
          });
          cursorY -= 20;
        }

        cursorY -= 8; // Paragraph gap
        if (cursorY < marginBottom + 20) break;
      }
    }
  }

  const pdfBytes = await doc.save();
  await fs.writeFile(outputPath, pdfBytes);
}

/* ==========================================================================
   Main Document Converter Dispatcher
   ========================================================================== */

export async function convertDocument(
  inputPath: string,
  outputPath: string,
  sourceFormat: string,
  targetFormat: string
): Promise<void> {
  const src = normalizeFormat(sourceFormat);
  const dst = normalizeFormat(targetFormat);
  const route = `${src}->${dst}`;

  switch (route) {
    // 1. Images to Office/PDF
    case 'png->pdf':
    case 'jpg->pdf':
    case 'webp->pdf':
    case 'avif->pdf':
      return convertImageToPdf(inputPath, outputPath, src);
    case 'png->docx':
    case 'jpg->docx':
    case 'webp->docx':
    case 'avif->docx':
      return convertImageToDocx(inputPath, outputPath, src);
    case 'png->pptx':
    case 'jpg->pptx':
    case 'webp->pptx':
    case 'avif->pptx':
      return convertImageToPptx(inputPath, outputPath, src);

    // TXT to PDF
    case 'txt->pdf':
      return convertTxtToPdf(inputPath, outputPath);

    // 2. Word (DOCX) Conversions
    case 'docx->pdf':
      return convertDocxToPdf(inputPath, outputPath);
    case 'docx->txt':
      return convertDocxToTxt(inputPath, outputPath);
    case 'docx->html':
      return convertDocxToHtml(inputPath, outputPath);

    // 3. PDF Conversions
    case 'pdf->docx':
      return convertPdfToDocx(inputPath, outputPath);
    case 'pdf->png':
      return convertPdfToImage(inputPath, outputPath, 'png');
    case 'pdf->jpg':
      return convertPdfToImage(inputPath, outputPath, 'jpg');
    case 'pdf->txt':
      return convertPdfToTxt(inputPath, outputPath);

    // 4. Spreadsheet (XLSX / CSV) Conversions
    case 'xlsx->pdf':
      return convertXlsxToPdf(inputPath, outputPath);
    case 'xlsx->csv':
      return convertXlsxToCsv(inputPath, outputPath);
    case 'xlsx->json':
      return convertXlsxToJson(inputPath, outputPath);
    case 'csv->xlsx':
      return convertCsvToXlsx(inputPath, outputPath);
    case 'json->xlsx':
      return convertJsonToXlsx(inputPath, outputPath);

    // 5. Presentation (PPTX) Conversions
    case 'pptx->pdf':
      return convertPptxToPdf(inputPath, outputPath);
    case 'pptx->txt':
      return convertPptxToTxt(inputPath, outputPath);

    default:
      throw new Error(`Unsupported document conversion route: ${sourceFormat} -> ${targetFormat}`);
  }
}
