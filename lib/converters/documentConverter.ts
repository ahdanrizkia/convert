import fs from 'node:fs/promises';
import sharp from 'sharp';
import { normalizeFormat } from './definitions';
import { convertImageToPdf, convertTxtToPdf } from './pdfConverter';
import type { PDFFont, PDFImage, PDFPage, RGB } from 'pdf-lib';

// Keep libvips memory flat
sharp.cache(false);

/**
 * Text sanitizer for standard WinAnsi font in pdf-lib (Helvetica)
 */
function sanitizeWinAnsiText(str: string, supportedSet?: Set<number>): string {
  const pre = str
    .replace(/\t/g, '    ')
    .replace(/\u00A0/g, ' ')
    .replace(/[\uF0B7\uF0A7\uF0D8\uF076\uF0FC]/g, '\u2022')
    .replace(/[\u2018\u2019\u2032]/g, "'")
    .replace(/[\u201C\u201D\u2033]/g, '"')
    .replace(/\u2014/g, '--')
    .replace(/[\u2013\u2212]/g, '-')
    .replace(/\u2026/g, '...')
    .replace(/\u2192/g, '->')
    .replace(/\u2190/g, '<-')
    .replace(/\u2194/g, '<->')
    .replace(/\u21D2/g, '=>')
    .replace(/[\u200B-\u200D\uFEFF]/g, '');

  let out = '';
  for (const ch of pre) {
    if (ch === '\n' || ch === '\r') {
      out += ch;
      continue;
    }
    const cp = ch.codePointAt(0);
    if (cp === undefined || cp < 0x20 || cp === 0x7f) continue;

    if (supportedSet) {
      if (supportedSet.has(cp)) {
        out += ch;
      } else {
        const decomposed = ch.normalize('NFKD');
        let fallback = '';
        for (const part of decomposed) {
          const pcp = part.codePointAt(0);
          if (pcp !== undefined && pcp >= 0x20 && pcp !== 0x7f && supportedSet.has(pcp)) {
            fallback += part;
          }
        }
        out += fallback.length > 0 ? fallback : cp === 0x2022 ? '-' : '?';
      }
    } else if ((cp >= 0x20 && cp <= 0x7e) || (cp >= 0xa0 && cp <= 0xff)) {
      out += ch;
    } else if (cp === 0x2022) {
      out += '-';
    } else {
      const decomposed = ch.normalize('NFKD').replace(/[^\x20-\x7E\xA0-\xFF]/g, '');
      out += decomposed.length > 0 ? decomposed : '?';
    }
  }
  return out;
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
   1. Images to Office/PDF (JPG, JFIF, PNG, WEBP, AVIF -> PDF, DOCX, PPTX)
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
                type: ext === 'jpg' || ext === 'jfif' ? 'jpg' : 'png',
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
  let mimeType =
    sourceFormat === 'jpg' || sourceFormat === 'jfif' ? 'image/jpeg' : 'image/png';

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
      Rich OpenXML Parser & Layout Engine (Tables, Lists, Styles, Images)
   ========================================================================== */

interface XmlNode {
  tag: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

const HTML_NAMED_ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
  middot: '\u00B7', // ·
  bull: '\u2022',   // •
  copy: '\u00A9',   // ©
  reg: '\u00AE',    // ®
  trade: '\u2122',  // ™
  ndash: '\u2013',  // –
  mdash: '\u2014',  // —
  lsquo: '\u2018',  // ‘
  rsquo: '\u2019',  // ’
  ldquo: '\u201C',  // “
  rdquo: '\u201D',  // ”
  larr: '\u2190',   // ←
  rarr: '\u2192',   // →
  uarr: '\u2191',   // ↑
  darr: '\u2193',   // ↓
  harr: '\u2194',   // ↔
  lArr: '\u21D0',   // ⇐
  rArr: '\u21D2',   // ⇒
  hArr: '\u21D4',   // ⇔
  hellip: '\u2026', // …
  times: '\u00D7',  // ×
  divide: '\u00F7', // ÷
  plusmn: '\u00B1', // ±
  deg: '\u00B0',    // °
  sup2: '\u00B2',   // ²
  sup3: '\u00B3',   // ³
  micro: '\u00B5',  // µ
  euro: '\u20AC',   // €
  pound: '\u00A3',  // £
  yen: '\u00A5',    // ¥
  cent: '\u00A2',   // ¢
};

function decodeXmlEntities(str: string): string {
  return str.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (full, entity: string) => {
    const named = HTML_NAMED_ENTITIES[entity.toLowerCase()];
    if (named !== undefined) return named;
    if (entity.startsWith('#x') || entity.startsWith('#X')) {
      const cp = parseInt(entity.slice(2), 16);
      return Number.isFinite(cp) && cp > 0 ? String.fromCodePoint(cp) : '';
    }
    if (entity.startsWith('#')) {
      const cp = parseInt(entity.slice(1), 10);
      return Number.isFinite(cp) && cp > 0 ? String.fromCodePoint(cp) : '';
    }
    return full;
  });
}

function parseXmlAttrs(rawAttrs: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (!rawAttrs) return attrs;
  const re = /([a-zA-Z0-9_:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(rawAttrs)) !== null) {
    const key = m[1];
    const val = decodeXmlEntities(m[2] ?? m[3] ?? '');
    attrs[key] = val;
    const shortKey = key.includes(':') ? key.slice(key.indexOf(':') + 1) : key;
    if (!(shortKey in attrs)) {
      attrs[shortKey] = val;
    }
  }
  return attrs;
}

function parseXml(xml: string): XmlNode {
  const cleaned = xml
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<\?[\s\S]*?\?>/g, '');

  const root: XmlNode = { tag: 'root', attrs: {}, children: [], text: '' };
  const stack: XmlNode[] = [root];

  const tokenRe = /<(\/?)([a-zA-Z0-9_:-]+)([^>]*?)(\/?)>|([^<]+)/g;
  let m: RegExpExecArray | null;

  while ((m = tokenRe.exec(cleaned)) !== null) {
    const isClosing = m[1] === '/';
    const tagName = m[2];
    const attrStr = m[3] || '';
    const isSelfClosing = m[4] === '/';
    const textChunk = m[5];

    if (textChunk !== undefined) {
      const top = stack[stack.length - 1];
      if (top) {
        top.text += decodeXmlEntities(textChunk);
      }
      continue;
    }

    if (isClosing) {
      for (let i = stack.length - 1; i >= 1; i--) {
        if (stack[i].tag === tagName) {
          stack.length = i;
          break;
        }
      }
      continue;
    }

    const node: XmlNode = {
      tag: tagName,
      attrs: parseXmlAttrs(attrStr),
      children: [],
      text: '',
    };

    const parent = stack[stack.length - 1];
    if (parent) {
      parent.children.push(node);
    }

    if (!isSelfClosing) {
      stack.push(node);
    }
  }

  return root;
}

function findChild(node: XmlNode | undefined, ...tags: string[]): XmlNode | undefined {
  if (!node) return undefined;
  for (const c of node.children) {
    if (tags.includes(c.tag)) return c;
  }
  return undefined;
}

function findChildren(node: XmlNode | undefined, ...tags: string[]): XmlNode[] {
  if (!node) return [];
  return node.children.filter((c) => tags.includes(c.tag));
}

function findDescendants(node: XmlNode | undefined, ...tags: string[]): XmlNode[] {
  if (!node) return [];
  const out: XmlNode[] = [];
  const visit = (cur: XmlNode) => {
    for (const c of cur.children) {
      if (tags.includes(c.tag)) {
        out.push(c);
      }
      visit(c);
    }
  };
  visit(node);
  return out;
}

function isXmlToggleEnabled(node: XmlNode | undefined): boolean | undefined {
  if (!node) return undefined;
  const val = (node.attrs['w:val'] ?? node.attrs.val ?? '').toLowerCase();
  if (val === '0' || val === 'false' || val === 'none' || val === 'off') return false;
  return true;
}

function parseHexRgb(hex: string | undefined): [number, number, number] | undefined {
  if (!hex) return undefined;
  const clean = hex.replace(/^#/, '').trim();
  if (!/^[0-9a-fA-F]{6}$/.test(clean)) return undefined;
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  return [r, g, b];
}

interface RunStyleProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number; // points
  colorRgb?: [number, number, number];
  vertAlign?: 'superscript' | 'subscript' | 'baseline';
}

interface ParaStyleProps {
  styleId?: string;
  name?: string;
  basedOn?: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  spaceBefore?: number; // points
  spaceAfter?: number; // points
  lineSpacingMult?: number;
  indentLeft?: number; // points
  indentRight?: number; // points
  indentFirstLine?: number; // points
  indentHanging?: number; // points
  numId?: string;
  ilvl?: number;
  pageBreakBefore?: boolean;
  shadingRgb?: [number, number, number];
  bottomBorder?: boolean;
  rPr: RunStyleProps;
}

function extractRunStyleProps(rPr: XmlNode | undefined): RunStyleProps {
  const out: RunStyleProps = {};
  if (!rPr) return out;

  const b = findChild(rPr, 'w:b', 'w:bCs');
  if (b) out.bold = isXmlToggleEnabled(b);

  const i = findChild(rPr, 'w:i', 'w:iCs');
  if (i) out.italic = isXmlToggleEnabled(i);

  const u = findChild(rPr, 'w:u');
  if (u) {
    const uVal = (u.attrs['w:val'] ?? u.attrs.val ?? 'single').toLowerCase();
    out.underline = uVal !== 'none' && uVal !== '0' && uVal !== 'false';
  }

  const strike = findChild(rPr, 'w:strike', 'w:dstrike');
  if (strike) out.strike = isXmlToggleEnabled(strike);

  const sz = findChild(rPr, 'w:sz', 'w:szCs');
  if (sz) {
    const halfPts = parseFloat(sz.attrs['w:val'] ?? sz.attrs.val ?? '');
    if (Number.isFinite(halfPts) && halfPts > 0) {
      out.fontSize = Math.min(Math.max(halfPts / 2, 6), 48);
    }
  }

  const color = findChild(rPr, 'w:color');
  if (color) {
    const cVal = color.attrs['w:val'] ?? color.attrs.val;
    if (cVal && cVal.toLowerCase() !== 'auto') {
      const parsed = parseHexRgb(cVal);
      if (parsed) out.colorRgb = parsed;
    }
  }

  const va = findChild(rPr, 'w:vertAlign');
  if (va) {
    const vVal = (va.attrs['w:val'] ?? va.attrs.val ?? '').toLowerCase();
    if (vVal === 'superscript' || vVal === 'subscript') {
      out.vertAlign = vVal;
    }
  }

  return out;
}

function mergeRunStyleProps(base: RunStyleProps, override: RunStyleProps): RunStyleProps {
  return {
    bold: override.bold ?? base.bold,
    italic: override.italic ?? base.italic,
    underline: override.underline ?? base.underline,
    strike: override.strike ?? base.strike,
    fontSize: override.fontSize ?? base.fontSize,
    colorRgb: override.colorRgb ?? base.colorRgb,
    vertAlign: override.vertAlign ?? base.vertAlign,
  };
}

function extractParaStyleProps(pPr: XmlNode | undefined): ParaStyleProps {
  const out: ParaStyleProps = { rPr: {} };
  if (!pPr) return out;

  const pStyle = findChild(pPr, 'w:pStyle');
  if (pStyle) {
    out.styleId = pStyle.attrs['w:val'] ?? pStyle.attrs.val;
  }

  const jc = findChild(pPr, 'w:jc');
  if (jc) {
    const jcVal = (jc.attrs['w:val'] ?? jc.attrs.val ?? '').toLowerCase();
    if (jcVal === 'center') out.align = 'center';
    else if (jcVal === 'right' || jcVal === 'end') out.align = 'right';
    else if (jcVal === 'both' || jcVal === 'distribute') out.align = 'justify';
    else out.align = 'left';
  }

  const spacing = findChild(pPr, 'w:spacing');
  if (spacing) {
    const beforeTwips = parseFloat(spacing.attrs['w:before'] ?? spacing.attrs.before ?? '');
    if (Number.isFinite(beforeTwips) && beforeTwips >= 0) {
      out.spaceBefore = Math.min(beforeTwips / 20, 48);
    }
    const afterTwips = parseFloat(spacing.attrs['w:after'] ?? spacing.attrs.after ?? '');
    if (Number.isFinite(afterTwips) && afterTwips >= 0) {
      out.spaceAfter = Math.min(afterTwips / 20, 48);
    }
    const lineTwips = parseFloat(spacing.attrs['w:line'] ?? spacing.attrs.line ?? '');
    const lineRule = (spacing.attrs['w:lineRule'] ?? spacing.attrs.lineRule ?? 'auto').toLowerCase();
    if (Number.isFinite(lineTwips) && lineTwips > 0 && lineRule === 'auto') {
      out.lineSpacingMult = Math.min(Math.max(lineTwips / 240, 1.0), 2.5);
    }
  }

  const ind = findChild(pPr, 'w:ind');
  if (ind) {
    const leftTwips = parseFloat(
      ind.attrs['w:left'] ?? ind.attrs.left ?? ind.attrs['w:start'] ?? ''
    );
    if (Number.isFinite(leftTwips)) {
      out.indentLeft = Math.max(0, Math.min(leftTwips / 20, 220));
    }
    const rightTwips = parseFloat(
      ind.attrs['w:right'] ?? ind.attrs.right ?? ind.attrs['w:end'] ?? ''
    );
    if (Number.isFinite(rightTwips)) {
      out.indentRight = Math.max(0, Math.min(rightTwips / 20, 220));
    }
    const hangingTwips = parseFloat(ind.attrs['w:hanging'] ?? ind.attrs.hanging ?? '');
    if (Number.isFinite(hangingTwips) && hangingTwips > 0) {
      out.indentHanging = Math.min(hangingTwips / 20, 120);
    }
    const firstLineTwips = parseFloat(ind.attrs['w:firstLine'] ?? ind.attrs.firstLine ?? '');
    if (Number.isFinite(firstLineTwips) && firstLineTwips > 0) {
      out.indentFirstLine = Math.min(firstLineTwips / 20, 120);
    }
  }

  const numPr = findChild(pPr, 'w:numPr');
  if (numPr) {
    const numIdNode = findChild(numPr, 'w:numId');
    const ilvlNode = findChild(numPr, 'w:ilvl');
    const numId = numIdNode?.attrs['w:val'] ?? numIdNode?.attrs.val;
    const ilvl = parseInt(ilvlNode?.attrs['w:val'] ?? ilvlNode?.attrs.val ?? '0', 10);
    if (numId && numId !== '0') {
      out.numId = numId;
      out.ilvl = Number.isFinite(ilvl) ? ilvl : 0;
    } else if (numId === '0') {
      out.numId = '0';
    }
  }

  const pbb = findChild(pPr, 'w:pageBreakBefore');
  if (pbb && isXmlToggleEnabled(pbb) !== false) {
    out.pageBreakBefore = true;
  }

  const shd = findChild(pPr, 'w:shd');
  if (shd) {
    const fill = shd.attrs['w:fill'] ?? shd.attrs.fill;
    if (fill && fill.toLowerCase() !== 'auto' && fill.toUpperCase() !== 'FFFFFF') {
      out.shadingRgb = parseHexRgb(fill);
    }
  }

  const pBdr = findChild(pPr, 'w:pBdr');
  if (pBdr) {
    const btm = findChild(pBdr, 'w:bottom');
    if (btm) {
      const bVal = (btm.attrs['w:val'] ?? btm.attrs.val ?? 'none').toLowerCase();
      if (bVal !== 'none' && bVal !== 'nil') {
        out.bottomBorder = true;
      }
    }
  }

  const pRPr = findChild(pPr, 'w:rPr');
  if (pRPr) {
    out.rPr = extractRunStyleProps(pRPr);
  }

  return out;
}

function mergeParaStyleProps(base: ParaStyleProps, override: ParaStyleProps): ParaStyleProps {
  return {
    styleId: override.styleId ?? base.styleId,
    name: override.name ?? base.name,
    basedOn: override.basedOn ?? base.basedOn,
    align: override.align ?? base.align,
    spaceBefore: override.spaceBefore ?? base.spaceBefore,
    spaceAfter: override.spaceAfter ?? base.spaceAfter,
    lineSpacingMult: override.lineSpacingMult ?? base.lineSpacingMult,
    indentLeft: override.indentLeft ?? base.indentLeft,
    indentRight: override.indentRight ?? base.indentRight,
    indentFirstLine: override.indentFirstLine ?? base.indentFirstLine,
    indentHanging: override.indentHanging ?? base.indentHanging,
    numId: override.numId ?? base.numId,
    ilvl: override.ilvl ?? base.ilvl,
    pageBreakBefore: override.pageBreakBefore ?? base.pageBreakBefore,
    shadingRgb: override.shadingRgb ?? base.shadingRgb,
    bottomBorder: override.bottomBorder ?? base.bottomBorder,
    rPr: mergeRunStyleProps(base.rPr, override.rPr),
  };
}

interface NumberingLevelDef {
  ilvl: number;
  start: number;
  numFmt: string;
  lvlText: string;
  indentLeft?: number;
  indentHanging?: number;
}

class DocxNumberingTracker {
  private abstractNums = new Map<string, Map<number, NumberingLevelDef>>();
  private numToAbstract = new Map<string, string>();
  private counters = new Map<string, number[]>();

  constructor(numberingXmlRoot?: XmlNode) {
    if (!numberingXmlRoot) return;
    const rootNode = findChild(numberingXmlRoot, 'w:numbering') ?? numberingXmlRoot;

    for (const absNode of findChildren(rootNode, 'w:abstractNum')) {
      const absId = absNode.attrs['w:abstractNumId'] ?? absNode.attrs.abstractNumId;
      if (!absId) continue;

      const lvlMap = new Map<number, NumberingLevelDef>();
      for (const lvlNode of findChildren(absNode, 'w:lvl')) {
        const ilvl = parseInt(lvlNode.attrs['w:ilvl'] ?? lvlNode.attrs.ilvl ?? '0', 10);
        const startNode = findChild(lvlNode, 'w:start');
        const start = parseInt(startNode?.attrs['w:val'] ?? startNode?.attrs.val ?? '1', 10);
        const fmtNode = findChild(lvlNode, 'w:numFmt');
        const numFmt = (fmtNode?.attrs['w:val'] ?? fmtNode?.attrs.val ?? 'decimal').toLowerCase();
        const txtNode = findChild(lvlNode, 'w:lvlText');
        const lvlText = txtNode?.attrs['w:val'] ?? txtNode?.attrs.val ?? '%1.';

        const pPr = findChild(lvlNode, 'w:pPr');
        const extracted = extractParaStyleProps(pPr);

        lvlMap.set(ilvl, {
          ilvl,
          start: Number.isFinite(start) ? start : 1,
          numFmt,
          lvlText,
          indentLeft: extracted.indentLeft,
          indentHanging: extracted.indentHanging,
        });
      }
      this.abstractNums.set(absId, lvlMap);
    }

    for (const numNode of findChildren(rootNode, 'w:num')) {
      const numId = numNode.attrs['w:numId'] ?? numNode.attrs.numId;
      const absIdNode = findChild(numNode, 'w:abstractNumId');
      const absId = absIdNode?.attrs['w:val'] ?? absIdNode?.attrs.val;
      if (numId && absId) {
        this.numToAbstract.set(numId, absId);
      }
    }
  }

  private formatCounter(val: number, fmt: string): string {
    switch (fmt) {
      case 'lowerletter': {
        const idx = Math.max(0, (val - 1) % 26);
        return String.fromCharCode(97 + idx);
      }
      case 'upperletter': {
        const idx = Math.max(0, (val - 1) % 26);
        return String.fromCharCode(65 + idx);
      }
      case 'lowerroman':
      case 'upperroman': {
        const romans: Array<[number, string]> = [
          [1000, 'm'],
          [900, 'cm'],
          [500, 'd'],
          [400, 'cd'],
          [100, 'c'],
          [90, 'xc'],
          [50, 'l'],
          [40, 'xl'],
          [10, 'x'],
          [9, 'ix'],
          [5, 'v'],
          [4, 'iv'],
          [1, 'i'],
        ];
        let rem = Math.max(1, Math.min(val, 3999));
        let out = '';
        for (const [n, sym] of romans) {
          while (rem >= n) {
            out += sym;
            rem -= n;
          }
        }
        return fmt === 'upperroman' ? out.toUpperCase() : out;
      }
      default:
        return String(val);
    }
  }

  nextMarker(
    numId: string,
    ilvl: number
  ): { prefix: string; indentLeft?: number; indentHanging?: number } {
    const absId = this.numToAbstract.get(numId);
    const lvlMap = absId ? this.abstractNums.get(absId) : undefined;
    const def = lvlMap?.get(ilvl) ?? {
      ilvl,
      start: 1,
      numFmt: 'decimal',
      lvlText: `%${ilvl + 1}.`,
      indentLeft: 18 * (ilvl + 1),
      indentHanging: 18,
    };

    if (def.numFmt === 'bullet') {
      return {
        prefix: '\u2022 ',
        indentLeft: def.indentLeft ?? 18 * (ilvl + 1),
        indentHanging: def.indentHanging ?? 14,
      };
    }

    if (def.numFmt === 'none') {
      return { prefix: '' };
    }

    let state = this.counters.get(numId);
    if (!state) {
      state = new Array(9).fill(0);
      this.counters.set(numId, state);
    }

    if (state[ilvl] === 0) {
      state[ilvl] = def.start;
    } else {
      state[ilvl] += 1;
    }

    for (let k = ilvl + 1; k < 9; k++) {
      state[k] = 0;
    }

    let formatted = def.lvlText || `%${ilvl + 1}.`;
    for (let k = 0; k <= ilvl; k++) {
      const kDef = lvlMap?.get(k) ?? def;
      const kVal = state[k] || kDef.start || 1;
      const strVal = this.formatCounter(kVal, kDef.numFmt);
      formatted = formatted.replace(new RegExp(`%${k + 1}`, 'g'), strVal);
    }

    return {
      prefix: `${formatted} `,
      indentLeft: def.indentLeft ?? 18 * (ilvl + 1),
      indentHanging: def.indentHanging ?? 18,
    };
  }
}

interface DocxImagePlacement {
  rId: string;
  widthPt: number;
  heightPt: number;
  isAnchor: boolean;
  alignH?: 'left' | 'center' | 'right';
  relFromH?: string;
  offsetHPt?: number;
  relFromV?: string;
  offsetVPt?: number;
  wrapMode?: 'square' | 'tight' | 'through' | 'topAndBottom' | 'none' | 'inline';
}

interface RichTextSpan {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  fontSize: number;
  color: [number, number, number];
  isPageBreak?: boolean;
  inlineImage?: DocxImagePlacement;
}

interface WrappedRichLine {
  tokens: RichTextSpan[];
  width: number;
  maxFontSize: number;
  lineHeight: number;
  isFirstLineOfPara: boolean;
}

async function convertDocxToPdf(inputPath: string, outputPath: string): Promise<void> {
  const JSZip = (await import('jszip')).default;
  const { PDFDocument, PageSizes, StandardFonts, rgb } = await import('pdf-lib');

  const zipBuffer = await fs.readFile(inputPath);
  const zip = await JSZip.loadAsync(zipBuffer);

  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('Invalid DOCX archive: missing word/document.xml');
  }

  const [docXmlStr, stylesXmlStr, numberingXmlStr, relsXmlStr] = await Promise.all([
    docXmlFile.async('text'),
    zip.file('word/styles.xml')?.async('text') ?? Promise.resolve(''),
    zip.file('word/numbering.xml')?.async('text') ?? Promise.resolve(''),
    zip.file('word/_rels/document.xml.rels')?.async('text') ?? Promise.resolve(''),
  ]);

  // 1. Parse Relationships (images & external hyperlinks)
  const rels = new Map<string, { target: string; external: boolean }>();
  if (relsXmlStr) {
    const relsRoot = parseXml(relsXmlStr);
    for (const relNode of findDescendants(relsRoot, 'Relationship')) {
      const id = relNode.attrs.Id;
      const target = relNode.attrs.Target;
      const external = (relNode.attrs.TargetMode ?? '').toLowerCase() === 'external';
      if (id && target) {
        rels.set(id, { target, external });
      }
    }
  }

  // 2. Parse Styles (`word/styles.xml`)
  let defaultParaStyle: ParaStyleProps = {
    align: 'left',
    spaceBefore: 0,
    spaceAfter: 6,
    lineSpacingMult: 1.15,
    rPr: {
      fontSize: 11,
      colorRgb: [0.1, 0.1, 0.1],
    },
  };

  const paraStyles = new Map<string, ParaStyleProps>();
  const charStyles = new Map<string, RunStyleProps>();

  if (stylesXmlStr) {
    const stylesRoot = parseXml(stylesXmlStr);
    const stylesContainer = findChild(stylesRoot, 'w:styles') ?? stylesRoot;

    const docDefaults = findChild(stylesContainer, 'w:docDefaults');
    if (docDefaults) {
      const rPrDefault = findChild(findChild(docDefaults, 'w:rPrDefault'), 'w:rPr');
      const pPrDefault = findChild(findChild(docDefaults, 'w:pPrDefault'), 'w:pPr');
      if (pPrDefault) {
        defaultParaStyle = mergeParaStyleProps(defaultParaStyle, extractParaStyleProps(pPrDefault));
      }
      if (rPrDefault) {
        defaultParaStyle.rPr = mergeRunStyleProps(
          defaultParaStyle.rPr,
          extractRunStyleProps(rPrDefault)
        );
      }
    }

    for (const stNode of findChildren(stylesContainer, 'w:style')) {
      const stType = (stNode.attrs['w:type'] ?? stNode.attrs.type ?? '').toLowerCase();
      const styleId = stNode.attrs['w:styleId'] ?? stNode.attrs.styleId;
      if (!styleId) continue;

      const nameNode = findChild(stNode, 'w:name');
      const stName = (nameNode?.attrs['w:val'] ?? nameNode?.attrs.val ?? '').toLowerCase();
      const basedOnNode = findChild(stNode, 'w:basedOn');
      const basedOn = basedOnNode?.attrs['w:val'] ?? basedOnNode?.attrs.val;

      const pPr = extractParaStyleProps(findChild(stNode, 'w:pPr'));
      const rPr = extractRunStyleProps(findChild(stNode, 'w:rPr'));

      // Built-in heading fallback sizes if theme/style leaves size implicit
      if (stName.startsWith('heading 1') || styleId.toLowerCase() === 'heading1') {
        if (!rPr.fontSize) rPr.fontSize = 16;
        if (rPr.bold === undefined) rPr.bold = true;
        if (pPr.spaceBefore === undefined) pPr.spaceBefore = 10;
        if (pPr.spaceAfter === undefined) pPr.spaceAfter = 4;
      } else if (stName.startsWith('heading 2') || styleId.toLowerCase() === 'heading2') {
        if (!rPr.fontSize) rPr.fontSize = 13.5;
        if (rPr.bold === undefined) rPr.bold = true;
        if (pPr.spaceBefore === undefined) pPr.spaceBefore = 8;
        if (pPr.spaceAfter === undefined) pPr.spaceAfter = 4;
      } else if (stName.startsWith('heading 3') || styleId.toLowerCase() === 'heading3') {
        if (!rPr.fontSize) rPr.fontSize = 12;
        if (rPr.bold === undefined) rPr.bold = true;
      } else if (stName === 'title' || styleId.toLowerCase() === 'title') {
        if (!rPr.fontSize) rPr.fontSize = 20;
        if (rPr.bold === undefined) rPr.bold = true;
      } else if (stName === 'subtitle' || styleId.toLowerCase() === 'subtitle') {
        if (!rPr.fontSize) rPr.fontSize = 13;
        if (rPr.italic === undefined) rPr.italic = true;
      }

      if (stType === 'character') {
        charStyles.set(styleId, rPr);
      } else {
        paraStyles.set(styleId, {
          ...pPr,
          styleId,
          name: stName,
          basedOn,
          rPr: mergeRunStyleProps(pPr.rPr, rPr),
        });
      }
    }
  }

  const resolveParaStyle = (styleId: string | undefined): ParaStyleProps => {
    if (!styleId) return defaultParaStyle;
    const chain: ParaStyleProps[] = [];
    const visited = new Set<string>();
    let curId: string | undefined = styleId;

    while (curId && !visited.has(curId)) {
      visited.add(curId);
      const st = paraStyles.get(curId);
      if (!st) break;
      chain.unshift(st);
      curId = st.basedOn;
    }

    let resolved = { ...defaultParaStyle, rPr: { ...defaultParaStyle.rPr } };
    for (const item of chain) {
      resolved = mergeParaStyleProps(resolved, item);
    }
    return resolved;
  };

  // 3. Parse Numbering (`word/numbering.xml`)
  const numberingTracker = new DocxNumberingTracker(
    numberingXmlStr ? parseXml(numberingXmlStr) : undefined
  );

  // 4. Parse Document Body (`word/document.xml`)
  const docRoot = parseXml(docXmlStr);
  const docNode = findChild(docRoot, 'w:document') ?? docRoot;
  const bodyNode = findChild(docNode, 'w:body') ?? docNode;

  // Check page dimensions & margins in <w:sectPr>
  let [pageW, pageH] = PageSizes.A4;
  let marginLeft = 64;
  let marginRight = 64;
  let marginTop = 60;
  let marginBottom = 60;

  const sectPr = findChild(bodyNode, 'w:sectPr') ?? findDescendants(bodyNode, 'w:sectPr')[0];
  if (sectPr) {
    const pgSz = findChild(sectPr, 'w:pgSz');
    if (pgSz) {
      const wTwips = parseFloat(pgSz.attrs['w:w'] ?? pgSz.attrs.w ?? '');
      const hTwips = parseFloat(pgSz.attrs['w:h'] ?? pgSz.attrs.h ?? '');
      if (Number.isFinite(wTwips) && wTwips > 2000 && Number.isFinite(hTwips) && hTwips > 2000) {
        pageW = wTwips / 20;
        pageH = hTwips / 20;
      }
    }
    const pgMar = findChild(sectPr, 'w:pgMar');
    if (pgMar) {
      const topTw = parseFloat(pgMar.attrs['w:top'] ?? pgMar.attrs.top ?? '');
      const btmTw = parseFloat(pgMar.attrs['w:bottom'] ?? pgMar.attrs.bottom ?? '');
      const leftTw = parseFloat(pgMar.attrs['w:left'] ?? pgMar.attrs.left ?? '');
      const rightTw = parseFloat(pgMar.attrs['w:right'] ?? pgMar.attrs.right ?? '');
      if (Number.isFinite(topTw)) marginTop = Math.min(Math.max(Math.abs(topTw) / 20, 36), 96);
      if (Number.isFinite(btmTw)) marginBottom = Math.min(Math.max(Math.abs(btmTw) / 20, 36), 96);
      if (Number.isFinite(leftTw)) marginLeft = Math.min(Math.max(Math.abs(leftTw) / 20, 36), 96);
      if (Number.isFinite(rightTw)) marginRight = Math.min(Math.max(Math.abs(rightTw) / 20, 36), 96);
    }
  }

  const maxContentW = pageW - marginLeft - marginRight;

  // 5. Initialize PDF Document and Font Family
  const pdfDoc = await PDFDocument.create();
  const [fontRegular, fontBold, fontItalic, fontBoldItalic] = await Promise.all([
    pdfDoc.embedFont(StandardFonts.Helvetica),
    pdfDoc.embedFont(StandardFonts.HelveticaBold),
    pdfDoc.embedFont(StandardFonts.HelveticaOblique),
    pdfDoc.embedFont(StandardFonts.HelveticaBoldOblique),
  ]);

  const supportedCharSet = new Set<number>(fontRegular.getCharacterSet());

  const selectFont = (bold?: boolean, italic?: boolean): PDFFont => {
    if (bold && italic) return fontBoldItalic;
    if (bold) return fontBold;
    if (italic) return fontItalic;
    return fontRegular;
  };

  // Cache embedded PDF images by relationship ID
  const embeddedImages = new Map<string, PDFImage | null>();

  const loadEmbeddedImage = async (rId: string): Promise<PDFImage | null> => {
    if (embeddedImages.has(rId)) {
      return embeddedImages.get(rId) ?? null;
    }

    const rel = rels.get(rId);
    if (!rel || rel.external) {
      embeddedImages.set(rId, null);
      return null;
    }

    const cleanTarget = rel.target.replace(/^\/+/, '').replace(/^\.\//, '');
    const candidates = [
      cleanTarget.startsWith('word/') ? cleanTarget : `word/${cleanTarget}`,
      cleanTarget,
    ];

    let zipEntry = null;
    for (const c of candidates) {
      const f = zip.file(c);
      if (f) {
        zipEntry = f;
        break;
      }
    }

    if (!zipEntry) {
      embeddedImages.set(rId, null);
      return null;
    }

    try {
      const rawImgBytes = await zipEntry.async('nodebuffer');
      const meta = await sharp(rawImgBytes, { failOn: 'error' }).metadata();
      let pdfImg: PDFImage;
      if (meta.hasAlpha) {
        const pngBuf = await sharp(rawImgBytes)
          .rotate()
          .png({ compressionLevel: 8 })
          .toBuffer();
        pdfImg = await pdfDoc.embedPng(pngBuf);
      } else {
        const jpgBuf = await sharp(rawImgBytes)
          .rotate()
          .flatten({ background: '#ffffff' })
          .jpeg({ quality: 92, mozjpeg: true, chromaSubsampling: '4:4:4' })
          .toBuffer();
        pdfImg = await pdfDoc.embedJpg(jpgBuf);
      }
      embeddedImages.set(rId, pdfImg);
      return pdfImg;
    } catch {
      embeddedImages.set(rId, null);
      return null;
    }
  };

  // Extract image placement details from <w:drawing> or <w:pict>
  const extractDrawingPlacements = (drawingNode: XmlNode): DocxImagePlacement[] => {
    const placements: DocxImagePlacement[] = [];

    // 1. Modern DrawingML (<wp:inline> / <wp:anchor>)
    const wpNodes = [
      ...findChildren(drawingNode, 'wp:inline', 'wp:anchor'),
      ...findDescendants(drawingNode, 'wp:inline', 'wp:anchor'),
    ];
    const seenWp = new Set<XmlNode>();

    for (const wp of wpNodes) {
      if (seenWp.has(wp)) continue;
      seenWp.add(wp);

      const isAnchor = wp.tag === 'wp:anchor';
      const blip = findDescendants(wp, 'a:blip')[0];
      const rId = blip?.attrs['r:embed'] ?? blip?.attrs.embed ?? blip?.attrs['r:link'];
      if (!rId) continue;

      const extent = findChild(wp, 'wp:extent') ?? findDescendants(wp, 'a:ext')[0];
      const cx = parseFloat(extent?.attrs.cx ?? '0');
      const cy = parseFloat(extent?.attrs.cy ?? '0');
      // 1 point = 12,700 EMUs
      let widthPt = cx > 0 ? cx / 12700 : 140;
      let heightPt = cy > 0 ? cy / 12700 : 100;

      if (widthPt > maxContentW) {
        const scale = maxContentW / widthPt;
        widthPt *= scale;
        heightPt *= scale;
      }

      let alignH: 'left' | 'center' | 'right' | undefined;
      let relFromH: string | undefined;
      let offsetHPt: number | undefined;
      let relFromV: string | undefined;
      let offsetVPt: number | undefined;
      let wrapMode: DocxImagePlacement['wrapMode'] = isAnchor ? 'square' : 'inline';

      if (isAnchor) {
        const posH = findChild(wp, 'wp:positionH');
        if (posH) {
          relFromH = (posH.attrs.relativeFrom ?? '').toLowerCase();
          const alignNode = findChild(posH, 'wp:align');
          const posOffsetNode = findChild(posH, 'wp:posOffset');
          const alignText = alignNode?.text.trim().toLowerCase();
          if (alignText === 'right' || alignText === 'center' || alignText === 'left') {
            alignH = alignText;
          }
          if (posOffsetNode) {
            const emuH = parseFloat(posOffsetNode.text.trim());
            if (Number.isFinite(emuH)) {
              offsetHPt = emuH / 12700;
            }
          }
        }

        const posV = findChild(wp, 'wp:positionV');
        if (posV) {
          relFromV = (posV.attrs.relativeFrom ?? '').toLowerCase();
          const posOffsetVNode = findChild(posV, 'wp:posOffset');
          if (posOffsetVNode) {
            const emuV = parseFloat(posOffsetVNode.text.trim());
            if (Number.isFinite(emuV)) {
              offsetVPt = emuV / 12700;
            }
          }
        }

        if (findChild(wp, 'wp:wrapTopAndBottom')) wrapMode = 'topAndBottom';
        else if (findChild(wp, 'wp:wrapNone')) wrapMode = 'none';
        else if (findChild(wp, 'wp:wrapTight')) wrapMode = 'tight';
        else if (findChild(wp, 'wp:wrapThrough')) wrapMode = 'through';
        else if (findChild(wp, 'wp:wrapSquare')) wrapMode = 'square';
      }

      placements.push({
        rId,
        widthPt,
        heightPt,
        isAnchor,
        alignH,
        relFromH,
        offsetHPt,
        relFromV,
        offsetVPt,
        wrapMode,
      });
    }

    // 2. Legacy VML (<v:imagedata>)
    if (placements.length === 0) {
      for (const vImg of findDescendants(drawingNode, 'v:imagedata')) {
        const rId = vImg.attrs['r:id'] ?? vImg.attrs.id;
        if (!rId) continue;
        placements.push({
          rId,
          widthPt: 150,
          heightPt: 110,
          isAnchor: false,
          wrapMode: 'inline',
        });
      }
    }

    return placements;
  };

  // Extract RichTextSpans and anchored images from a <w:p> node
  const extractParagraphContent = (
    pNode: XmlNode,
    inTable = false
  ): {
    style: ParaStyleProps;
    spans: RichTextSpan[];
    anchoredImages: DocxImagePlacement[];
  } => {
    const pPrNode = findChild(pNode, 'w:pPr');
    const inlinePStyle = extractParaStyleProps(pPrNode);
    const baseStyle = resolveParaStyle(inlinePStyle.styleId);
    const style = mergeParaStyleProps(baseStyle, inlinePStyle);

    const spans: RichTextSpan[] = [];
    const anchoredImages: DocxImagePlacement[] = [];

    const collectRunsFromContainer = (
      container: XmlNode,
      inheritedRunStyle: RunStyleProps
    ) => {
      for (const child of container.children) {
        if (child.tag === 'w:pPr') continue;

        if (child.tag === 'w:hyperlink') {
          const linkStyle: RunStyleProps = {
            ...inheritedRunStyle,
            colorRgb: [0.02, 0.38, 0.76],
            underline: true,
          };
          collectRunsFromContainer(child, linkStyle);
          continue;
        }

        if (
          child.tag === 'w:ins' ||
          child.tag === 'w:smartTag' ||
          child.tag === 'w:sdt' ||
          child.tag === 'w:sdtContent' ||
          child.tag === 'w:fldSimple'
        ) {
          collectRunsFromContainer(child, inheritedRunStyle);
          continue;
        }

        if (child.tag === 'w:r') {
          const rPrNode = findChild(child, 'w:rPr');
          const rStyleNode = findChild(rPrNode, 'w:rStyle');
          const rStyleId = rStyleNode?.attrs['w:val'] ?? rStyleNode?.attrs.val;
          let charStyle = rStyleId ? (charStyles.get(rStyleId) ?? {}) : {};
          if (rStyleId && rStyleId.toLowerCase().includes('hyperlink')) {
            charStyle = {
              ...charStyle,
              colorRgb: charStyle.colorRgb ?? [0.02, 0.38, 0.76],
              underline: charStyle.underline ?? true,
            };
          }

          const runRPr = extractRunStyleProps(rPrNode);
          const mergedRun = mergeRunStyleProps(
            mergeRunStyleProps(inheritedRunStyle, charStyle),
            runRPr
          );

          let fontSize = mergedRun.fontSize ?? 11;
          if (mergedRun.vertAlign === 'superscript' || mergedRun.vertAlign === 'subscript') {
            fontSize = Math.max(7, fontSize * 0.75);
          }

          const baseSpanProps = {
            bold: Boolean(mergedRun.bold),
            italic: Boolean(mergedRun.italic),
            underline: Boolean(mergedRun.underline),
            strike: Boolean(mergedRun.strike),
            fontSize,
            color: mergedRun.colorRgb ?? ([0.1, 0.1, 0.1] as [number, number, number]),
          };

          for (const rChild of child.children) {
            if (rChild.tag === 'w:t') {
              const sanitized = sanitizeWinAnsiText(rChild.text, supportedCharSet);
              if (sanitized.length > 0) {
                spans.push({
                  ...baseSpanProps,
                  text: sanitized,
                });
              }
            } else if (rChild.tag === 'w:tab') {
              spans.push({
                ...baseSpanProps,
                text: '    ',
              });
            } else if (rChild.tag === 'w:br' || rChild.tag === 'w:cr') {
              const brType = (rChild.attrs['w:type'] ?? rChild.attrs.type ?? '').toLowerCase();
              if (brType === 'page' && !inTable) {
                spans.push({
                  ...baseSpanProps,
                  text: '',
                  isPageBreak: true,
                });
              } else {
                spans.push({
                  ...baseSpanProps,
                  text: '\n',
                });
              }
            } else if (rChild.tag === 'w:drawing' || rChild.tag === 'w:pict') {
              const imgs = extractDrawingPlacements(rChild);
              for (const img of imgs) {
                if (img.isAnchor && !inTable) {
                  anchoredImages.push(img);
                } else {
                  spans.push({
                    ...baseSpanProps,
                    text: '',
                    inlineImage: img,
                  });
                }
              }
            }
          }
        }
      }
    };

    collectRunsFromContainer(pNode, style.rPr);

    // Handle paragraph numbering (<w:numPr>) if paragraph has visible content
    const hasVisibleText = spans.some((s) => s.text.trim().length > 0);
    if (style.numId && style.numId !== '0' && hasVisibleText) {
      const marker = numberingTracker.nextMarker(style.numId, style.ilvl ?? 0);
      if (marker.prefix) {
        const firstTextSpan = spans.find((s) => s.text.trim().length > 0);
        const markerFontSize = firstTextSpan?.fontSize ?? style.rPr.fontSize ?? 11;
        const markerBold = firstTextSpan?.bold ?? Boolean(style.rPr.bold);
        const markerItalic = firstTextSpan?.italic ?? Boolean(style.rPr.italic);
        const markerColor =
          firstTextSpan?.color ?? style.rPr.colorRgb ?? ([0.1, 0.1, 0.1] as [number, number, number]);

        spans.unshift({
          text: sanitizeWinAnsiText(marker.prefix, supportedCharSet),
          bold: markerBold,
          italic: markerItalic,
          underline: false,
          strike: false,
          fontSize: markerFontSize,
          color: markerColor,
        });

        if (style.indentLeft === undefined && marker.indentLeft !== undefined) {
          style.indentLeft = marker.indentLeft;
        }
        if (style.indentHanging === undefined && marker.indentHanging !== undefined) {
          style.indentHanging = marker.indentHanging;
        }
      }
    }

    return { style, spans, anchoredImages };
  };

  // Word-wrap RichTextSpans into lines fitting within maxWidth
  const wrapRichSpans = (
    spans: RichTextSpan[],
    firstLineMaxW: number,
    otherLinesMaxW: number,
    lineSpacingMult = 1.15,
    defaultFontSize = 11
  ): WrappedRichLine[] => {
    const lines: WrappedRichLine[] = [];
    let curTokens: RichTextSpan[] = [];
    let curWidth = 0;
    let isFirstLine = true;

    const getLineMaxW = () => Math.max(24, isFirstLine ? firstLineMaxW : otherLinesMaxW);

    const flushLine = (nextIsNewParagraphLine = false) => {
      // Trim trailing whitespace tokens
      while (curTokens.length > 0 && /^\s+$/.test(curTokens[curTokens.length - 1].text)) {
        const removed = curTokens.pop()!;
        const f = selectFont(removed.bold, removed.italic);
        curWidth = Math.max(0, curWidth - f.widthOfTextAtSize(removed.text, removed.fontSize));
      }

      let maxFontSize = defaultFontSize;
      for (const t of curTokens) {
        if (t.fontSize > maxFontSize) maxFontSize = t.fontSize;
      }

      lines.push({
        tokens: curTokens,
        width: curWidth,
        maxFontSize,
        lineHeight: Math.max(12, maxFontSize * Math.max(1.15, lineSpacingMult)),
        isFirstLineOfPara: isFirstLine,
      });

      curTokens = [];
      curWidth = 0;
      isFirstLine = nextIsNewParagraphLine;
    };

    const measureToken = (span: RichTextSpan, txt: string): number => {
      const f = selectFont(span.bold, span.italic);
      return f.widthOfTextAtSize(txt, span.fontSize);
    };

    for (const span of spans) {
      if (span.isPageBreak || span.inlineImage) continue;
      if (!span.text) continue;

      const subLines = span.text.split(/\r?\n/);
      for (let sIdx = 0; sIdx < subLines.length; sIdx++) {
        if (sIdx > 0) {
          flushLine(false);
        }

        const segment = subLines[sIdx];
        if (!segment) continue;

        const pieces = segment.match(/\S+|\s+/g) ?? [segment];
        for (const piece of pieces) {
          const isSpace = /^\s+$/.test(piece);
          if (curTokens.length === 0 && isSpace && !isFirstLine) {
            continue;
          }

          const pieceW = measureToken(span, piece);
          const availW = getLineMaxW();

          if (curWidth + pieceW <= availW) {
            curTokens.push({ ...span, text: piece });
            curWidth += pieceW;
            continue;
          }

          if (isSpace) {
            flushLine(false);
            continue;
          }

          if (curTokens.length > 0) {
            flushLine(false);
          }

          const freshAvailW = getLineMaxW();
          if (pieceW <= freshAvailW) {
            curTokens.push({ ...span, text: piece });
            curWidth = pieceW;
          } else {
            // Break long unbroken token character by character
            let chunk = '';
            let chunkW = 0;
            for (const ch of piece) {
              const chW = measureToken(span, ch);
              if (chunk.length > 0 && chunkW + chW > getLineMaxW()) {
                curTokens.push({ ...span, text: chunk });
                curWidth = chunkW;
                flushLine(false);
                chunk = ch;
                chunkW = chW;
              } else {
                chunk += ch;
                chunkW += chW;
              }
            }
            if (chunk.length > 0) {
              curTokens.push({ ...span, text: chunk });
              curWidth = chunkW;
            }
          }
        }
      }
    }

    if (curTokens.length > 0 || lines.length === 0) {
      flushLine(false);
    }

    return lines;
  };

  // Draw a single WrappedRichLine onto a PDF page
  const drawWrappedLine = (
    targetPage: PDFPage,
    line: WrappedRichLine,
    xStart: number,
    availWidth: number,
    baselineY: number,
    align: 'left' | 'center' | 'right' | 'justify' = 'left'
  ) => {
    if (line.tokens.length === 0) return;

    let offsetX = xStart;
    if (align === 'center') {
      offsetX = xStart + Math.max(0, (availWidth - line.width) / 2);
    } else if (align === 'right') {
      offsetX = xStart + Math.max(0, availWidth - line.width);
    }

    for (const tok of line.tokens) {
      if (!tok.text) continue;
      const font = selectFont(tok.bold, tok.italic);
      const tokW = font.widthOfTextAtSize(tok.text, tok.fontSize);
      const tokColor: RGB = rgb(tok.color[0], tok.color[1], tok.color[2]);

      if (tok.text.trim().length > 0) {
        targetPage.drawText(tok.text, {
          x: offsetX,
          y: baselineY,
          size: tok.fontSize,
          font,
          color: tokColor,
        });
      }

      if (tok.underline && (tok.text.trim().length > 0 || tokW > 0)) {
        targetPage.drawLine({
          start: { x: offsetX, y: baselineY - 1.5 },
          end: { x: offsetX + tokW, y: baselineY - 1.5 },
          thickness: Math.max(0.5, tok.fontSize * 0.065),
          color: tokColor,
        });
      }

      if (tok.strike && (tok.text.trim().length > 0 || tokW > 0)) {
        targetPage.drawLine({
          start: { x: offsetX, y: baselineY + tok.fontSize * 0.3 },
          end: { x: offsetX + tokW, y: baselineY + tok.fontSize * 0.3 },
          thickness: Math.max(0.5, tok.fontSize * 0.06),
          color: tokColor,
        });
      }

      offsetX += tokW;
    }
  };

  // Pagination & floating image state
  let page = pdfDoc.addPage([pageW, pageH]);
  let cursorY = pageH - marginTop;

  interface ActiveFloatBox {
    side: 'left' | 'right';
    xLeft: number;
    xRight: number;
    bottomY: number;
  }

  let activeFloat: ActiveFloatBox | null = null;

  const startNewPage = () => {
    page = pdfDoc.addPage([pageW, pageH]);
    cursorY = pageH - marginTop;
    activeFloat = null;
  };

  // Draw an image placement (anchored or inline)
  const renderImagePlacement = async (
    imgSpec: DocxImagePlacement,
    paraAlign: 'left' | 'center' | 'right' | 'justify' = 'left'
  ) => {
    const pdfImg = await loadEmbeddedImage(imgSpec.rId);
    if (!pdfImg) return;

    const w = Math.min(imgSpec.widthPt, maxContentW);
    const h = imgSpec.heightPt * (w / Math.max(1, imgSpec.widthPt));

    if (
      imgSpec.isAnchor &&
      imgSpec.wrapMode !== 'topAndBottom' &&
      w <= maxContentW * 0.62
    ) {
      // Determine horizontal placement of anchored floating image
      let drawX = marginLeft;
      let side: 'left' | 'right' = 'right';

      if (imgSpec.alignH === 'right') {
        drawX = pageW - marginRight - w;
        side = 'right';
      } else if (imgSpec.offsetHPt !== undefined) {
        const baseH =
          imgSpec.relFromH === 'page' || imgSpec.relFromH === 'leftmargin'
            ? 0
            : marginLeft;
        drawX = Math.min(
          pageW - marginRight - w,
          Math.max(marginLeft, baseH + imgSpec.offsetHPt)
        );
        side = drawX - marginLeft >= maxContentW * 0.3 ? 'right' : 'left';
      } else {
        drawX = pageW - marginRight - w;
        side = 'right';
      }

      let topY = cursorY;
      if (imgSpec.offsetVPt !== undefined) {
        if (imgSpec.relFromV === 'page') {
          topY = Math.min(pageH - 18, Math.max(cursorY - 60, pageH - imgSpec.offsetVPt));
        } else if (imgSpec.relFromV === 'margin' || imgSpec.relFromV === 'topmargin') {
          topY = Math.min(
            pageH - 18,
            Math.max(cursorY - 60, pageH - marginTop - imgSpec.offsetVPt)
          );
        } else {
          const clampedOffset = Math.max(-24, Math.min(imgSpec.offsetVPt, 60));
          topY = Math.min(pageH - 18, cursorY - clampedOffset);
        }
      }
      if (topY - h < marginBottom) {
        startNewPage();
        topY = cursorY;
      }

      page.drawImage(pdfImg, {
        x: drawX,
        y: topY - h,
        width: w,
        height: h,
      });

      activeFloat = {
        side,
        xLeft: drawX,
        xRight: drawX + w,
        bottomY: topY - h - 4,
      };
      return;
    }

    // Inline or block-level image
    if (activeFloat && cursorY > activeFloat.bottomY) {
      cursorY = activeFloat.bottomY;
      activeFloat = null;
    }

    if (cursorY - h < marginBottom && cursorY < pageH - marginTop - 10) {
      startNewPage();
    }

    let drawX = marginLeft;
    const align = imgSpec.alignH ?? paraAlign;
    if (align === 'center') {
      drawX = marginLeft + Math.max(0, (maxContentW - w) / 2);
    } else if (align === 'right') {
      drawX = pageW - marginRight - w;
    }

    page.drawImage(pdfImg, {
      x: drawX,
      y: cursorY - h,
      width: w,
      height: h,
    });

    cursorY -= h + 8;
  };

  // Render a top-level paragraph <w:p>
  const renderParagraph = async (pNode: XmlNode) => {
    const { style, spans, anchoredImages } = extractParagraphContent(pNode, false);

    if (style.pageBreakBefore && cursorY < pageH - marginTop - 10) {
      startNewPage();
    }

    // Render any anchored images attached to this paragraph first so text wraps around them
    for (const anc of anchoredImages) {
      await renderImagePlacement(anc, style.align);
    }

    // Check for explicit inline page breaks or inline images
    for (const sp of spans) {
      if (sp.isPageBreak) {
        startNewPage();
      } else if (sp.inlineImage) {
        await renderImagePlacement(sp.inlineImage, style.align);
      }
    }

    const textSpans = spans.filter((s) => !s.isPageBreak && !s.inlineImage && s.text.length > 0);
    if (textSpans.length === 0) {
      if (anchoredImages.length === 0 && !spans.some((s) => s.inlineImage)) {
        const emptyFontSize = style.rPr.fontSize ?? 11;
        const emptyLineH = Math.max(12, emptyFontSize * (style.lineSpacingMult ?? 1.15));
        cursorY -= emptyLineH + (style.spaceAfter ?? 4);
      }
      return;
    }

    const spaceBefore = style.spaceBefore ?? 0;
    const spaceAfter = style.spaceAfter ?? 5;
    const lineMult = style.lineSpacingMult ?? 1.15;
    const baseFontSize = style.rPr.fontSize ?? 11;

    if (spaceBefore > 0 && cursorY < pageH - marginTop - 4) {
      cursorY -= spaceBefore;
    }

    // Compute available horizontal bounds taking into account indentation and active floating image
    const indentLeft = style.indentLeft ?? 0;
    const indentRight = style.indentRight ?? 0;
    const hanging = style.indentHanging ?? 0;
    const firstLineIndent = style.indentFirstLine ?? 0;

    let baseLeftX = marginLeft + indentLeft;
    let baseRightX = pageW - marginRight - indentRight;

    if (activeFloat && cursorY > activeFloat.bottomY) {
      if (activeFloat.side === 'right') {
        baseRightX = Math.min(baseRightX, activeFloat.xLeft - 10);
      } else {
        baseLeftX = Math.max(baseLeftX, activeFloat.xRight + 10);
      }
    } else if (activeFloat && cursorY <= activeFloat.bottomY) {
      activeFloat = null;
    }

    const firstLineLeftX = Math.max(marginLeft, baseLeftX - hanging + firstLineIndent);
    const otherLinesLeftX = baseLeftX;

    const firstLineAvailW = Math.max(40, baseRightX - firstLineLeftX);
    const otherLinesAvailW = Math.max(40, baseRightX - otherLinesLeftX);

    const wrappedLines = wrapRichSpans(
      textSpans,
      firstLineAvailW,
      otherLinesAvailW,
      lineMult,
      baseFontSize
    );

    for (const line of wrappedLines) {
      if (cursorY - line.lineHeight < marginBottom) {
        startNewPage();
      }

      if (activeFloat && cursorY <= activeFloat.bottomY) {
        activeFloat = null;
      }

      const lineX = line.isFirstLineOfPara ? firstLineLeftX : otherLinesLeftX;
      const lineW = line.isFirstLineOfPara ? firstLineAvailW : otherLinesAvailW;
      const baselineY = cursorY - line.maxFontSize;

      drawWrappedLine(page, line, lineX, lineW, baselineY, style.align ?? 'left');
      cursorY -= line.lineHeight;
    }

    if (style.bottomBorder) {
      page.drawLine({
        start: { x: baseLeftX, y: cursorY + 2 },
        end: { x: baseRightX, y: cursorY + 2 },
        thickness: 0.75,
        color: rgb(0.5, 0.5, 0.5),
      });
      cursorY -= 4;
    }

    cursorY -= spaceAfter;
  };

  // Render a Word table <w:tbl>
  const renderTable = async (tblNode: XmlNode) => {
    // If there is an active floating image, clear past it if the table is wide
    if (activeFloat && cursorY > activeFloat.bottomY) {
      cursorY = Math.min(cursorY, activeFloat.bottomY - 4);
      activeFloat = null;
    }

    const trNodes = findChildren(tblNode, 'w:tr');
    if (trNodes.length === 0) return;

    // 1. Determine grid column widths from <w:tblGrid> or first row cells
    const tblGrid = findChild(tblNode, 'w:tblGrid');
    const gridColNodes = findChildren(tblGrid, 'w:gridCol');
    let colWidthsPt: number[] = [];

    for (const gc of gridColNodes) {
      const wTwips = parseFloat(gc.attrs['w:w'] ?? gc.attrs.w ?? '');
      if (Number.isFinite(wTwips) && wTwips > 0) {
        colWidthsPt.push(wTwips / 20);
      }
    }

    // Determine max column count across rows
    let numCols = colWidthsPt.length;
    for (const tr of trNodes) {
      let rowCols = 0;
      for (const tc of findChildren(tr, 'w:tc')) {
        const tcPr = findChild(tc, 'w:tcPr');
        const gs = findChild(tcPr, 'w:gridSpan');
        const span = Math.max(1, parseInt(gs?.attrs['w:val'] ?? gs?.attrs.val ?? '1', 10) || 1);
        rowCols += span;
      }
      if (rowCols > numCols) numCols = rowCols;
    }

    if (numCols === 0) return;

    // Fill or infer missing column widths from <w:tcW> or content length
    if (colWidthsPt.length < numCols) {
      const inferred = new Array<number>(numCols).fill(0);
      for (let i = 0; i < colWidthsPt.length; i++) {
        inferred[i] = colWidthsPt[i];
      }

      for (const tr of trNodes) {
        let colCursor = 0;
        for (const tc of findChildren(tr, 'w:tc')) {
          const tcPr = findChild(tc, 'w:tcPr');
          const gs = findChild(tcPr, 'w:gridSpan');
          const span = Math.max(1, parseInt(gs?.attrs['w:val'] ?? gs?.attrs.val ?? '1', 10) || 1);
          const tcW = findChild(tcPr, 'w:tcW');
          const wVal = parseFloat(tcW?.attrs['w:w'] ?? tcW?.attrs.w ?? '');
          const wType = (tcW?.attrs['w:type'] ?? tcW?.attrs.type ?? 'dxa').toLowerCase();

          if (span === 1 && colCursor < numCols && inferred[colCursor] === 0) {
            if (Number.isFinite(wVal) && wVal > 0) {
              if (wType === 'dxa') {
                inferred[colCursor] = wVal / 20;
              } else if (wType === 'pct') {
                inferred[colCursor] = (wVal / 5000) * maxContentW;
              }
            }
          }
          colCursor += span;
        }
      }

      for (let c = 0; c < numCols; c++) {
        if (inferred[c] <= 0) {
          inferred[c] = maxContentW / numCols;
        }
      }
      colWidthsPt = inferred;
    }

    // Scale columns if total width exceeds maxContentW, or enforce min column width
    let totalTableW = colWidthsPt.reduce((acc, w) => acc + w, 0);
    if (totalTableW > maxContentW || totalTableW <= 0) {
      const scale = totalTableW > 0 ? maxContentW / totalTableW : 1;
      colWidthsPt = colWidthsPt.map((w) => Math.max(28, w * scale));
      totalTableW = colWidthsPt.reduce((acc, w) => acc + w, 0);
      if (totalTableW > maxContentW) {
        const fixScale = maxContentW / totalTableW;
        colWidthsPt = colWidthsPt.map((w) => w * fixScale);
        totalTableW = maxContentW;
      }
    }

    // Table alignment
    const tblPr = findChild(tblNode, 'w:tblPr');
    const tblJc = findChild(tblPr, 'w:jc');
    const tblAlign = (tblJc?.attrs['w:val'] ?? tblJc?.attrs.val ?? 'left').toLowerCase();
    let tableStartX = marginLeft;
    if (tblAlign === 'center') {
      tableStartX = marginLeft + Math.max(0, (maxContentW - totalTableW) / 2);
    } else if (tblAlign === 'right' || tblAlign === 'end') {
      tableStartX = pageW - marginRight - totalTableW;
    }

    // Check if table borders are explicitly hidden
    let drawBorders = true;
    const tblBorders = findChild(tblPr, 'w:tblBorders');
    if (tblBorders) {
      const borderSides = findChildren(
        tblBorders,
        'w:top',
        'w:bottom',
        'w:left',
        'w:right',
        'w:insideH',
        'w:insideV'
      );
      if (
        borderSides.length > 0 &&
        borderSides.every((b) => {
          const v = (b.attrs['w:val'] ?? b.attrs.val ?? '').toLowerCase();
          return v === 'none' || v === 'nil';
        })
      ) {
        drawBorders = false;
      }
    }

    const cellPadX = 5;
    const cellPadTop = 4;
    const cellPadBottom = 4;

    if (cursorY < pageH - marginTop - 6) {
      cursorY -= 4;
    }

    // 2. Render each row
    for (const tr of trNodes) {
      const tcNodes = findChildren(tr, 'w:tc');
      if (tcNodes.length === 0) continue;

      interface PreparedCell {
        x: number;
        width: number;
        fillRgb?: [number, number, number];
        paragraphs: Array<{
          lines: WrappedRichLine[];
          align: 'left' | 'center' | 'right' | 'justify';
          spaceAfter: number;
        }>;
        contentHeight: number;
      }

      const preparedCells: PreparedCell[] = [];
      let colIdx = 0;
      let curX = tableStartX;
      let maxCellHeight = 18;

      for (const tc of tcNodes) {
        if (colIdx >= numCols) break;
        const tcPr = findChild(tc, 'w:tcPr');
        const gs = findChild(tcPr, 'w:gridSpan');
        const span = Math.max(
          1,
          Math.min(numCols - colIdx, parseInt(gs?.attrs['w:val'] ?? gs?.attrs.val ?? '1', 10) || 1)
        );

        let cellW = 0;
        for (let k = 0; k < span; k++) {
          cellW += colWidthsPt[colIdx + k] ?? 60;
        }

        let fillRgb: [number, number, number] | undefined;
        const shd = findChild(tcPr, 'w:shd');
        if (shd) {
          const fillHex = shd.attrs['w:fill'] ?? shd.attrs.fill;
          if (fillHex && fillHex.toLowerCase() !== 'auto' && fillHex.toUpperCase() !== 'FFFFFF') {
            fillRgb = parseHexRgb(fillHex);
          }
        }

        const innerAvailW = Math.max(16, cellW - cellPadX * 2);
        const cellPList = findChildren(tc, 'w:p');
        const cellParas: PreparedCell['paragraphs'] = [];
        let cellContentH = cellPadTop + cellPadBottom;

        for (let pIdx = 0; pIdx < cellPList.length; pIdx++) {
          const { style, spans } = extractParagraphContent(cellPList[pIdx], true);
          const lines = wrapRichSpans(
            spans,
            innerAvailW,
            innerAvailW,
            style.lineSpacingMult ?? 1.12,
            style.rPr.fontSize ?? 10.5
          );
          const spaceAfter = pIdx < cellPList.length - 1 ? Math.min(style.spaceAfter ?? 3, 6) : 0;

          for (const ln of lines) {
            cellContentH += ln.lineHeight;
          }
          cellContentH += spaceAfter;

          cellParas.push({
            lines,
            align: style.align ?? 'left',
            spaceAfter,
          });
        }

        if (cellContentH > maxCellHeight) {
          maxCellHeight = cellContentH;
        }

        preparedCells.push({
          x: curX,
          width: cellW,
          fillRgb,
          paragraphs: cellParas,
          contentHeight: cellContentH,
        });

        curX += cellW;
        colIdx += span;
      }

      // Page break if row does not fit on current page
      if (cursorY - maxCellHeight < marginBottom && cursorY < pageH - marginTop - 10) {
        startNewPage();
      }

      const rowTopY = cursorY;
      const rowBottomY = cursorY - maxCellHeight;

      for (const cell of preparedCells) {
        if (cell.fillRgb) {
          page.drawRectangle({
            x: cell.x,
            y: rowBottomY,
            width: cell.width,
            height: maxCellHeight,
            color: rgb(cell.fillRgb[0], cell.fillRgb[1], cell.fillRgb[2]),
          });
        }

        if (drawBorders) {
          page.drawRectangle({
            x: cell.x,
            y: rowBottomY,
            width: cell.width,
            height: maxCellHeight,
            borderWidth: 0.5,
            borderColor: rgb(0.65, 0.65, 0.65),
          });
        }

        let cellCursorY = rowTopY - cellPadTop;
        const innerW = Math.max(16, cell.width - cellPadX * 2);

        for (const para of cell.paragraphs) {
          for (const ln of para.lines) {
            const baselineY = cellCursorY - ln.maxFontSize;
            if (baselineY >= rowBottomY - 2) {
              drawWrappedLine(page, ln, cell.x + cellPadX, innerW, baselineY, para.align);
            }
            cellCursorY -= ln.lineHeight;
          }
          cellCursorY -= para.spaceAfter;
        }
      }

      cursorY -= maxCellHeight;
    }

    cursorY -= 8;
  };

  // Traverse <w:body> elements in document order
  const renderBodyChildren = async (container: XmlNode) => {
    for (const child of container.children) {
      if (child.tag === 'w:p') {
        await renderParagraph(child);
      } else if (child.tag === 'w:tbl') {
        await renderTable(child);
      } else if (child.tag === 'w:sdt') {
        const sdtContent = findChild(child, 'w:sdtContent');
        if (sdtContent) {
          await renderBodyChildren(sdtContent);
        }
      }
    }
  };

  await renderBodyChildren(bodyNode);

  const pdfBytes = await pdfDoc.save();
  await fs.writeFile(outputPath, pdfBytes);
}

async function convertDocxToPptx(inputPath: string, outputPath: string): Promise<void> {
  const tempPdfPath = `${inputPath}.tmp-${Date.now()}.pdf`;
  try {
    await convertDocxToPdf(inputPath, tempPdfPath);
    await convertPdfToPptx(tempPdfPath, outputPath);
  } finally {
    await fs.unlink(tempPdfPath).catch(() => {});
  }
}

async function convertDocxToTxt(inputPath: string, outputPath: string): Promise<void> {
  const JSZip = (await import('jszip')).default;
  const zipBuffer = await fs.readFile(inputPath);
  const zip = await JSZip.loadAsync(zipBuffer);
  const docXmlFile = zip.file('word/document.xml');

  if (!docXmlFile) {
    const mammoth = await import('mammoth');
    const { value } = await mammoth.extractRawText({ path: inputPath });
    await fs.writeFile(outputPath, value || '', 'utf8');
    return;
  }

  const [docXmlStr, numberingXmlStr] = await Promise.all([
    docXmlFile.async('text'),
    zip.file('word/numbering.xml')?.async('text') ?? Promise.resolve(''),
  ]);

  const numberingTracker = new DocxNumberingTracker(
    numberingXmlStr ? parseXml(numberingXmlStr) : undefined
  );
  const docRoot = parseXml(docXmlStr);
  const bodyNode =
    findChild(findChild(docRoot, 'w:document') ?? docRoot, 'w:body') ?? docRoot;

  const extractParaPlainText = (pNode: XmlNode): string => {
    const pPr = extractParaStyleProps(findChild(pNode, 'w:pPr'));
    let text = '';
    const walk = (n: XmlNode) => {
      for (const c of n.children) {
        if (c.tag === 'w:t') text += c.text;
        else if (c.tag === 'w:tab') text += '\t';
        else if (c.tag === 'w:br' || c.tag === 'w:cr') text += '\n';
        else if (c.tag !== 'w:pPr') walk(c);
      }
    };
    walk(pNode);

    if (pPr.numId && pPr.numId !== '0' && text.trim().length > 0) {
      const marker = numberingTracker.nextMarker(pPr.numId, pPr.ilvl ?? 0);
      if (marker.prefix) {
        text = `${marker.prefix}${text}`;
      }
    }
    return text;
  };

  const outLines: string[] = [];
  const walkBody = (container: XmlNode) => {
    for (const child of container.children) {
      if (child.tag === 'w:p') {
        outLines.push(extractParaPlainText(child));
      } else if (child.tag === 'w:tbl') {
        for (const tr of findChildren(child, 'w:tr')) {
          const cells = findChildren(tr, 'w:tc').map((tc) =>
            findChildren(tc, 'w:p')
              .map((p) => extractParaPlainText(p).replace(/\r?\n/g, ' '))
              .join(' ')
              .trim()
          );
          outLines.push(cells.join('\t'));
        }
        outLines.push('');
      } else if (child.tag === 'w:sdt') {
        const sdtContent = findChild(child, 'w:sdtContent');
        if (sdtContent) walkBody(sdtContent);
      }
    }
  };

  walkBody(bodyNode);
  await fs.writeFile(outputPath, outLines.join('\n'), 'utf8');
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
    img { max-width: 100%; height: auto; }
  </style>
</head>
<body>
${value || '<p>(Empty Document)</p>'}
</body>
</html>`;

  await fs.writeFile(outputPath, html, 'utf8');
}

/* ==========================================================================
   3. PDF Conversions (PDF -> DOCX, PNG, JPG, JFIF, TXT)
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

async function convertPdfToPptx(inputPath: string, outputPath: string): Promise<void> {
  const { PDFParse } = await import('pdf-parse');
  const pptxgen = (await import('pptxgenjs')).default;

  const pdfBuffer = await fs.readFile(inputPath);
  const parser = new PDFParse({ data: new Uint8Array(pdfBuffer) });

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';

  try {
    const screenshots = await parser.getScreenshot({ scale: 2.5 });
    const pages = screenshots?.pages ?? [];

    if (pages.length === 0) {
      const parseResult = await parser.getText();
      const rawText = parseResult?.text || '';
      const lines = rawText.split(/\r?\n/).filter((l) => l.trim().length > 0);
      const chunkSize = 8;
      for (let i = 0; i < Math.max(1, lines.length); i += chunkSize) {
        const slide = pres.addSlide();
        const slideLines = lines.slice(i, i + chunkSize);
        slide.addText(`Page ${Math.floor(i / chunkSize) + 1}`, {
          x: 0.8,
          y: 0.5,
          w: 8.4,
          h: 0.7,
          fontSize: 20,
          bold: true,
          color: '1E293B',
        });
        slide.addText(slideLines.join('\n'), {
          x: 0.8,
          y: 1.4,
          w: 8.4,
          h: 3.6,
          fontSize: 13,
          color: '334155',
        });
      }
    } else {
      for (const pg of pages) {
        const slide = pres.addSlide();
        const ratio = (pg.width || 800) / (pg.height || 600);

        // Slide dimension is 10.0 x 5.625 inches
        let w = 9.2;
        let h = w / ratio;
        if (h > 5.1) {
          h = 5.1;
          w = h * ratio;
        }
        const x = (10.0 - w) / 2;
        const y = (5.625 - h) / 2;

        slide.addImage({
          data: `data:image/png;base64,${Buffer.from(pg.data).toString('base64')}`,
          x: Number(x.toFixed(2)),
          y: Number(y.toFixed(2)),
          w: Number(w.toFixed(2)),
          h: Number(h.toFixed(2)),
        });
      }
    }
  } finally {
    await parser.destroy();
    delete (globalThis as unknown as { pdfjsLib?: unknown }).pdfjsLib;
  }

  const buffer = (await pres.write({ outputType: 'nodebuffer' })) as Buffer;
  await fs.writeFile(outputPath, new Uint8Array(buffer));
}

async function convertPdfToImage(
  inputPath: string,
  outputPath: string,
  targetFormat: 'png' | 'jpg' | 'jfif'
): Promise<void> {
  const { PDFParse } = await import('pdf-parse');
  const pdfBuffer = await fs.readFile(inputPath);
  const parser = new PDFParse({ data: new Uint8Array(pdfBuffer) });

  let pngBytes: Uint8Array | null = null;
  try {
    const screenshot = await parser.getScreenshot({ pageNumber: 1, scale: 3.0 } as any);
    if (screenshot?.pages?.[0]?.data) {
      pngBytes = screenshot.pages[0].data;
    }
  } finally {
    await parser.destroy();
  }

  if (!pngBytes) {
    throw new Error('Failed to render PDF page 1 to image.');
  }

  if (targetFormat === 'jpg' || targetFormat === 'jfif') {
    const jpgBytes = await sharp(pngBytes)
      .flatten({ background: '#ffffff' })
      .jpeg({ quality: 94, mozjpeg: true, chromaSubsampling: '4:4:4' })
      .toBuffer();
    await fs.writeFile(outputPath, new Uint8Array(jpgBytes));
  } else {
    const optimizedPng = await sharp(pngBytes)
      .png({ compressionLevel: 8 })
      .toBuffer();
    await fs.writeFile(outputPath, new Uint8Array(optimizedPng));
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
   6. Additional Document Conversions (MD, HTML, EPUB, XLSX -> PPTX, PPTX -> Image)
   ========================================================================== */

async function convertTxtToDocx(inputPath: string, outputPath: string): Promise<void> {
  const { Document, Packer, Paragraph, TextRun } = await import('docx');
  const text = await fs.readFile(inputPath, 'utf8');
  const lines = text.split(/\r?\n/);
  const children = lines.map(
    (l) => new Paragraph({ children: [new TextRun(l)], spacing: { after: 80 } })
  );
  const doc = new Document({
    sections: [{ children: children.length > 0 ? children : [new Paragraph('(Empty File)')] }],
  });
  const buf = await Packer.toBuffer(doc);
  await fs.writeFile(outputPath, new Uint8Array(buf));
}

async function convertTxtToPptx(inputPath: string, outputPath: string): Promise<void> {
  const tempPdf = `${inputPath}.tmp-${Date.now()}.pdf`;
  try {
    await convertTxtToPdf(inputPath, tempPdf);
    await convertPdfToPptx(tempPdf, outputPath);
  } finally {
    await fs.unlink(tempPdf).catch(() => {});
  }
}

async function convertPptxToImage(
  inputPath: string,
  outputPath: string,
  targetFormat: 'png' | 'jpg'
): Promise<void> {
  const tempPdf = `${inputPath}.tmp-${Date.now()}.pdf`;
  try {
    await convertPptxToPdf(inputPath, tempPdf);
    await convertPdfToImage(tempPdf, outputPath, targetFormat);
  } finally {
    await fs.unlink(tempPdf).catch(() => {});
  }
}

async function convertXlsxToPptx(inputPath: string, outputPath: string): Promise<void> {
  const XLSX = await import('xlsx');
  const pptxgen = (await import('pptxgenjs')).default;
  const wb = XLSX.readFile(inputPath);
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const rawRows: unknown[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (!rawRows || rawRows.length === 0) continue;

    const chunkSize = 9;
    for (let i = 0; i < rawRows.length; i += chunkSize) {
      const chunk = rawRows.slice(i, i + chunkSize);
      const slide = pres.addSlide();
      slide.addText(`Sheet: ${sheetName}${i > 0 ? ' (Cont.)' : ''}`, {
        x: 0.8,
        y: 0.4,
        w: 8.4,
        h: 0.6,
        fontSize: 20,
        bold: true,
        color: '0F172A',
      });

      const tableData = chunk.map((r, rIdx) => {
        const isHeader = i === 0 && rIdx === 0;
        return (Array.isArray(r) ? r : []).map((c) => ({
          text: c !== undefined && c !== null ? String(c) : '',
          options: {
            bold: isHeader,
            fill: isHeader ? '1E293B' : rIdx % 2 === 1 ? 'F8FAFC' : 'FFFFFF',
            color: isHeader ? 'FFFFFF' : '1E293B',
            fontSize: isHeader ? 10.5 : 9.5,
            margin: [4, 6, 4, 6],
          },
        }));
      });

      slide.addTable(tableData as any, {
        x: 0.8,
        y: 1.2,
        w: 8.4,
        border: { pt: 0.5, color: 'CBD5E1' },
      });
    }
  }

  const buf = (await pres.write({ outputType: 'nodebuffer' })) as Buffer;
  await fs.writeFile(outputPath, new Uint8Array(buf));
}

async function convertMdToHtml(inputPath: string, outputPath: string): Promise<void> {
  const { marked } = await import('marked');
  const md = await fs.readFile(inputPath, 'utf8');
  const htmlBody = await marked.parse(md);

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Converted Markdown</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; max-width: 800px; margin: 2rem auto; padding: 1.5rem; line-height: 1.6; color: #1a1a1a; background: #fafafa; }
    p { margin-bottom: 1em; }
    h1, h2, h3 { color: #0f172a; margin-top: 1.5em; margin-bottom: 0.5em; }
    code { background: #f1f5f9; padding: 0.2em 0.4em; border-radius: 4px; font-family: monospace; font-size: 0.9em; }
    pre { background: #0f172a; color: #f8fafc; padding: 1rem; border-radius: 6px; overflow-x: auto; }
    pre code { background: none; color: inherit; padding: 0; }
    table { border-collapse: collapse; width: 100%; margin: 1em 0; }
    td, th { border: 1px solid #cbd5e1; padding: 6px 10px; text-align: left; }
    th { background: #f8fafc; }
    blockquote { border-left: 4px solid #94a3b8; margin: 1em 0; padding-left: 1rem; color: #475569; }
  </style>
</head>
<body>
${htmlBody}
</body>
</html>`;

  await fs.writeFile(outputPath, html, 'utf8');
}

async function convertMdToTxt(inputPath: string, outputPath: string): Promise<void> {
  const md = await fs.readFile(inputPath, 'utf8');
  const plain = md
    .replace(/!\[.*?\]\(.*?\)/g, '')
    .replace(/\[(.*?)\]\(.*?\)/g, '$1')
    .replace(/^#+\s+/gm, '')
    .replace(/[*_~`]/g, '');
  await fs.writeFile(outputPath, plain, 'utf8');
}

async function convertMdToDocx(inputPath: string, outputPath: string): Promise<void> {
  const { marked } = await import('marked');
  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    HeadingLevel,
    Table,
    TableRow,
    TableCell,
    WidthType,
  } = await import('docx');
  const md = await fs.readFile(inputPath, 'utf8');
  const tokens = marked.lexer(md);
  const children: (InstanceType<typeof Paragraph> | InstanceType<typeof Table>)[] = [];

  const parseInlineTokens = (
    subTokens: any[] | undefined,
    baseProps: { bold?: boolean; italic?: boolean } = {}
  ): InstanceType<typeof TextRun>[] => {
    const runs: InstanceType<typeof TextRun>[] = [];
    if (!subTokens || !Array.isArray(subTokens)) return runs;

    for (const t of subTokens) {
      if (t.type === 'strong') {
        const nested = parseInlineTokens(t.tokens, { ...baseProps, bold: true });
        if (nested.length > 0) runs.push(...nested);
        else runs.push(new TextRun({ text: t.text || '', bold: true }));
      } else if (t.type === 'em') {
        const nested = parseInlineTokens(t.tokens, { ...baseProps, italic: true });
        if (nested.length > 0) runs.push(...nested);
        else runs.push(new TextRun({ text: t.text || '', italics: true }));
      } else if (t.type === 'codespan') {
        runs.push(
          new TextRun({
            text: t.text || '',
            font: 'Courier New',
            shading: { fill: 'F1F5F9' },
            color: '0F172A',
          })
        );
      } else if (t.type === 'del') {
        runs.push(new TextRun({ text: t.text || '', strike: true }));
      } else if (t.type === 'link') {
        runs.push(new TextRun({ text: t.text || '', color: '0284C7', underline: {} }));
      } else if (t.type === 'text') {
        runs.push(
          new TextRun({
            text: t.text || '',
            bold: baseProps.bold,
            italics: baseProps.italic,
          })
        );
      } else if (t.text) {
        runs.push(
          new TextRun({
            text: t.text,
            bold: baseProps.bold,
            italics: baseProps.italic,
          })
        );
      }
    }
    return runs;
  };

  for (const tok of tokens) {
    if (tok.type === 'heading') {
      const heading =
        tok.depth === 1
          ? HeadingLevel.HEADING_1
          : tok.depth === 2
          ? HeadingLevel.HEADING_2
          : HeadingLevel.HEADING_3;
      const runs = parseInlineTokens((tok as any).tokens);
      children.push(
        new Paragraph({
          children: runs.length > 0 ? runs : [new TextRun(tok.text)],
          heading,
          spacing: { before: 240, after: 120 },
        })
      );
    } else if (tok.type === 'paragraph') {
      const runs = parseInlineTokens((tok as any).tokens);
      children.push(
        new Paragraph({
          children: runs.length > 0 ? runs : [new TextRun(tok.text)],
          spacing: { after: 120 },
        })
      );
    } else if (tok.type === 'list') {
      for (const item of (tok as any).items || []) {
        const runs = parseInlineTokens(item.tokens);
        children.push(
          new Paragraph({
            children: runs.length > 0 ? runs : [new TextRun(item.text)],
            bullet: { level: 0 },
            spacing: { after: 60 },
          })
        );
      }
    } else if (tok.type === 'code') {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: tok.text, font: 'Courier New' })],
          spacing: { after: 120 },
        })
      );
    } else if (tok.type === 'table') {
      const tableTok = tok as any;
      const headerCells = (tableTok.header || []).map(
        (h: any) =>
          new TableCell({
            children: [
              new Paragraph({
                children: parseInlineTokens(h.tokens, { bold: true }),
                spacing: { before: 60, after: 60 },
              }),
            ],
            shading: { fill: 'F1F5F9' },
          })
      );
      const rowList: InstanceType<typeof TableRow>[] = [
        new TableRow({
          tableHeader: true,
          children: headerCells,
        }),
      ];
      for (const row of tableTok.rows || []) {
        rowList.push(
          new TableRow({
            children: (row || []).map(
              (cell: any) =>
                new TableCell({
                  children: [
                    new Paragraph({
                      children: parseInlineTokens(cell.tokens),
                      spacing: { before: 60, after: 60 },
                    }),
                  ],
                })
            ),
          })
        );
      }
      children.push(
        new Table({
          rows: rowList,
          width: { size: 100, type: WidthType.PERCENTAGE },
        })
      );
    }
  }

  if (children.length === 0) children.push(new Paragraph('(Empty Document)'));
  const doc = new Document({ sections: [{ children }] });
  const buf = await Packer.toBuffer(doc);
  await fs.writeFile(outputPath, new Uint8Array(buf));
}

async function convertMdToPdf(inputPath: string, outputPath: string): Promise<void> {
  const tempDocx = `${inputPath}.tmp-${Date.now()}.docx`;
  try {
    await convertMdToDocx(inputPath, tempDocx);
    await convertDocxToPdf(tempDocx, outputPath);
  } finally {
    await fs.unlink(tempDocx).catch(() => {});
  }
}

async function convertHtmlToTxt(inputPath: string, outputPath: string): Promise<void> {
  const html = await fs.readFile(inputPath, 'utf8');
  const plain = decodeXmlEntities(
    html
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|h[1-6]|li|header|nav|section|article|footer|tr|blockquote)>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
  )
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n\n')
    .trim();
  await fs.writeFile(outputPath, plain, 'utf8');
}

async function convertHtmlToDocx(inputPath: string, outputPath: string): Promise<void> {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel } = await import('docx');
  const rawHtml = await fs.readFile(inputPath, 'utf8');

  const processed = rawHtml
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<h1[^>]*>/gi, '\n[[H1]]')
    .replace(/<h2[^>]*>/gi, '\n[[H2]]')
    .replace(/<h3[^>]*>/gi, '\n[[H3]]')
    .replace(/<h[4-6][^>]*>/gi, '\n[[H4]]')
    .replace(/<li[^>]*>/gi, '\n[[LI]]')
    .replace(/<blockquote[^>]*>/gi, '\n[[BQ]]')
    .replace(/<p[^>]*>/gi, '\n[[P]]')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(h[1-6]|p|li|blockquote|div|header|nav|section|article|footer|main|tr|table)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ');

  const rawLines = processed.split(/\r?\n/);
  const children: InstanceType<typeof Paragraph>[] = [];

  for (const rawLine of rawLines) {
    let line = decodeXmlEntities(rawLine).replace(/[ \t]+/g, ' ').trim();
    if (!line) continue;

    let heading: (typeof HeadingLevel)[keyof typeof HeadingLevel] | undefined;
    let isBullet = false;
    let isBlockquote = false;

    if (line.startsWith('[[H1]]')) {
      heading = HeadingLevel.HEADING_1;
      line = line.slice(6).trim();
    } else if (line.startsWith('[[H2]]')) {
      heading = HeadingLevel.HEADING_2;
      line = line.slice(6).trim();
    } else if (line.startsWith('[[H3]]')) {
      heading = HeadingLevel.HEADING_3;
      line = line.slice(6).trim();
    } else if (line.startsWith('[[H4]]')) {
      heading = HeadingLevel.HEADING_4;
      line = line.slice(6).trim();
    } else if (line.startsWith('[[LI]]')) {
      isBullet = true;
      line = line.slice(6).trim();
    } else if (line.startsWith('[[BQ]]')) {
      isBlockquote = true;
      line = line.slice(6).trim();
    } else if (line.startsWith('[[P]]')) {
      line = line.slice(5).trim();
    }

    if (!line) continue;

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: line,
            italics: isBlockquote,
            bold: Boolean(heading),
          }),
        ],
        heading,
        bullet: isBullet ? { level: 0 } : undefined,
        spacing: { before: heading ? 200 : 0, after: heading ? 100 : 80 },
      })
    );
  }

  if (children.length === 0) {
    children.push(new Paragraph('(Empty Document)'));
  }

  const doc = new Document({ sections: [{ children }] });
  const buf = await Packer.toBuffer(doc);
  await fs.writeFile(outputPath, new Uint8Array(buf));
}

async function convertHtmlToPdf(inputPath: string, outputPath: string): Promise<void> {
  const tempDocx = `${inputPath}.tmp-${Date.now()}.docx`;
  try {
    await convertHtmlToDocx(inputPath, tempDocx);
    await convertDocxToPdf(tempDocx, outputPath);
  } finally {
    await fs.unlink(tempDocx).catch(() => {});
  }
}

async function convertEpubToTxt(inputPath: string, outputPath: string): Promise<void> {
  const JSZip = (await import('jszip')).default;
  const zipBuf = await fs.readFile(inputPath);
  const zip = await JSZip.loadAsync(zipBuf);

  const htmlFiles = Object.keys(zip.files)
    .filter((n) => /\.(x?html|xml|htm)$/i.test(n) && !n.includes('toc') && !n.includes('nav'))
    .sort();

  const texts: string[] = [];
  for (const fn of htmlFiles) {
    const raw = await zip.files[fn].async('text');
    const plain = raw
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (plain) texts.push(plain);
  }

  await fs.writeFile(outputPath, texts.join('\n\n') || '(Empty EPUB Document)', 'utf8');
}

async function convertEpubToPdf(inputPath: string, outputPath: string): Promise<void> {
  const JSZip = (await import('jszip')).default;
  const zipBuf = await fs.readFile(inputPath);
  const zip = await JSZip.loadAsync(zipBuf);

  const htmlFiles = Object.keys(zip.files)
    .filter((n) => /\.(x?html|xml|htm)$/i.test(n) && !n.includes('toc') && !n.includes('nav'))
    .sort();

  const htmlChunks: string[] = [];
  for (const fn of htmlFiles) {
    const raw = await zip.files[fn].async('text');
    htmlChunks.push(raw);
  }

  const tempHtml = `${inputPath}.tmp-${Date.now()}.html`;
  try {
    await fs.writeFile(tempHtml, htmlChunks.join('\n<hr/>\n'), 'utf8');
    await convertHtmlToPdf(tempHtml, outputPath);
  } finally {
    await fs.unlink(tempHtml).catch(() => {});
  }
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
    case 'jfif->pdf':
    case 'webp->pdf':
    case 'avif->pdf':
    case 'svg->pdf':
    case 'tiff->pdf':
    case 'gif->pdf':
    case 'bmp->pdf':
    case 'ico->pdf':
      return convertImageToPdf(inputPath, outputPath, src);
    case 'png->docx':
    case 'jpg->docx':
    case 'jfif->docx':
    case 'webp->docx':
    case 'avif->docx':
      return convertImageToDocx(inputPath, outputPath, src);
    case 'png->pptx':
    case 'jpg->pptx':
    case 'jfif->pptx':
    case 'webp->pptx':
    case 'avif->pptx':
      return convertImageToPptx(inputPath, outputPath, src);

    // TXT to Office/PDF
    case 'txt->pdf':
      return convertTxtToPdf(inputPath, outputPath);
    case 'txt->docx':
      return convertTxtToDocx(inputPath, outputPath);
    case 'txt->pptx':
      return convertTxtToPptx(inputPath, outputPath);

    // 2. Word (DOCX) Conversions
    case 'docx->pdf':
      return convertDocxToPdf(inputPath, outputPath);
    case 'docx->pptx':
      return convertDocxToPptx(inputPath, outputPath);
    case 'docx->txt':
      return convertDocxToTxt(inputPath, outputPath);
    case 'docx->html':
      return convertDocxToHtml(inputPath, outputPath);

    // 3. PDF Conversions
    case 'pdf->docx':
      return convertPdfToDocx(inputPath, outputPath);
    case 'pdf->pptx':
      return convertPdfToPptx(inputPath, outputPath);
    case 'pdf->png':
      return convertPdfToImage(inputPath, outputPath, 'png');
    case 'pdf->jpg':
      return convertPdfToImage(inputPath, outputPath, 'jpg');
    case 'pdf->jfif':
      return convertPdfToImage(inputPath, outputPath, 'jfif');
    case 'pdf->txt':
      return convertPdfToTxt(inputPath, outputPath);

    // 4. Spreadsheet (XLSX / CSV) Conversions
    case 'xlsx->pdf':
      return convertXlsxToPdf(inputPath, outputPath);
    case 'xlsx->pptx':
      return convertXlsxToPptx(inputPath, outputPath);
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
    case 'pptx->png':
      return convertPptxToImage(inputPath, outputPath, 'png');
    case 'pptx->jpg':
      return convertPptxToImage(inputPath, outputPath, 'jpg');
    case 'pptx->txt':
      return convertPptxToTxt(inputPath, outputPath);

    // 6. Markdown (MD) Conversions
    case 'md->html':
      return convertMdToHtml(inputPath, outputPath);
    case 'md->txt':
      return convertMdToTxt(inputPath, outputPath);
    case 'md->docx':
      return convertMdToDocx(inputPath, outputPath);
    case 'md->pdf':
      return convertMdToPdf(inputPath, outputPath);

    // 7. HTML Conversions
    case 'html->txt':
      return convertHtmlToTxt(inputPath, outputPath);
    case 'html->docx':
      return convertHtmlToDocx(inputPath, outputPath);
    case 'html->pdf':
      return convertHtmlToPdf(inputPath, outputPath);

    // 8. EPUB Conversions
    case 'epub->txt':
      return convertEpubToTxt(inputPath, outputPath);
    case 'epub->pdf':
      return convertEpubToPdf(inputPath, outputPath);

    default:
      throw new Error(`Unsupported document conversion route: ${sourceFormat} -> ${targetFormat}`);
  }
}
