import { createReadStream, createWriteStream, type WriteStream } from 'node:fs';
import { once } from 'node:events';
import readline from 'node:readline';
import type { DataFormat } from './definitions';

const SUPPORTED_DATA_FORMATS = new Set<string>(['json', 'csv', 'txt']);

export function isDataFormat(fmt: string): fmt is DataFormat {
  return SUPPORTED_DATA_FORMATS.has(fmt);
}

/**
 * Backpressure-safe write helper for Node WriteStreams.
 */
async function writeChunk(ws: WriteStream, chunk: string): Promise<void> {
  if (!ws.write(chunk, 'utf8')) {
    await once(ws, 'drain');
  }
}

/**
 * Cleanly closes a WriteStream and awaits disk flush completion.
 */
async function closeWriteStream(ws: WriteStream): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    ws.on('error', reject);
    ws.end(() => resolve());
  });
}

/**
 * Escapes a single cell value per RFC 4180 CSV specification.
 */
function escapeCsvCell(value: unknown): string {
  if (value === null || value === undefined) return '';
  const str =
    typeof value === 'object' ? JSON.stringify(value) : String(value);
  if (
    str.includes(',') ||
    str.includes('"') ||
    str.includes('\n') ||
    str.includes('\r')
  ) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

/**
 * Formats an arbitrary parsed JSON value into a normalized record object.
 */
function toRecord(item: unknown): Record<string, unknown> {
  if (item !== null && typeof item === 'object' && !Array.isArray(item)) {
    return item as Record<string, unknown>;
  }
  if (Array.isArray(item)) {
    const rec: Record<string, unknown> = {};
    for (let i = 0; i < item.length; i++) {
      rec[`col_${i + 1}`] = item[i];
    }
    return rec;
  }
  return { value: item };
}

/**
 * Streaming top-level JSON generator.
 * Yields individual elements of a top-level JSON array `[ {...}, {...} ]` one by one
 * without loading a multi-megabyte JSON DOM into memory at once, or yields a single
 * root object/primitive if the root is not an array.
 */
async function* streamJsonItems(inputPath: string): AsyncGenerator<unknown> {
  const rs = createReadStream(inputPath, {
    encoding: 'utf8',
    highWaterMark: 64 * 1024,
  });

  let started = false;
  let rootIsArray = false;
  let rootFinished = false;
  let depth = 0;
  let inString = false;
  let escaped = false;
  let tokenBuffer = '';
  let yieldedCount = 0;

  for await (let chunk of rs) {
    if (!started) {
      // Strip optional UTF-8 BOM
      if (chunk.charCodeAt(0) === 0xfeff) {
        chunk = chunk.slice(1);
      }
    }

    for (let i = 0; i < chunk.length; i++) {
      const ch = chunk[i];

      if (!started) {
        if (/\s/.test(ch)) continue;
        started = true;
        if (ch === '[') {
          rootIsArray = true;
          depth = 1;
          continue;
        } else {
          rootIsArray = false;
          tokenBuffer += ch;
          continue;
        }
      }

      if (!rootIsArray) {
        tokenBuffer += ch;
        continue;
      }

      if (rootFinished) {
        if (!/\s/.test(ch)) {
          throw new Error('Invalid trailing characters after root JSON array.');
        }
        continue;
      }

      if (inString) {
        tokenBuffer += ch;
        if (escaped) {
          escaped = false;
        } else if (ch === '\\') {
          escaped = true;
        } else if (ch === '"') {
          inString = false;
        }
        continue;
      }

      if (ch === '"') {
        inString = true;
        tokenBuffer += ch;
        continue;
      }

      if (ch === '{' || ch === '[') {
        depth++;
        tokenBuffer += ch;
        continue;
      }

      if (ch === '}' || ch === ']') {
        depth--;
        if (depth === 0 && ch === ']') {
          const trimmed = tokenBuffer.trim();
          if (trimmed.length > 0) {
            yield JSON.parse(trimmed);
            yieldedCount++;
          }
          tokenBuffer = '';
          rootFinished = true;
          continue;
        }
        if (depth < 0) {
          throw new Error('Malformed JSON bracket structure.');
        }
        tokenBuffer += ch;
        continue;
      }

      if (ch === ',' && depth === 1) {
        const trimmed = tokenBuffer.trim();
        if (trimmed.length === 0) {
          throw new Error('Unexpected comma in JSON array.');
        }
        yield JSON.parse(trimmed);
        yieldedCount++;
        tokenBuffer = '';
        continue;
      }

      tokenBuffer += ch;
    }
  }

  if (!started) {
    throw new Error('Uploaded JSON file is empty.');
  }

  if (!rootIsArray) {
    const trimmed = tokenBuffer.trim();
    const parsed = JSON.parse(trimmed);
    yield parsed;
    return;
  }

  if (!rootFinished) {
    throw new Error('Unexpected end of JSON array.');
  }

  void yieldedCount;
}

/**
 * Streaming RFC 4180 CSV parser generator.
 * Correctly handles quoted fields containing commas, escaped double quotes (`""`),
 * embedded CR/LF newlines inside quotes, and UTF-8 BOM.
 */
async function* streamCsvRows(inputPath: string): AsyncGenerator<string[]> {
  const rs = createReadStream(inputPath, {
    encoding: 'utf8',
    highWaterMark: 64 * 1024,
  });

  let firstChunk = true;
  let inQuotes = false;
  let quotePendingClose = false;
  let currentCell = '';
  let currentRow: string[] = [];

  for await (let chunk of rs) {
    if (firstChunk) {
      firstChunk = false;
      if (chunk.charCodeAt(0) === 0xfeff) {
        chunk = chunk.slice(1);
      }
    }

    for (let i = 0; i < chunk.length; i++) {
      const ch = chunk[i];

      if (quotePendingClose) {
        if (ch === '"') {
          // Escaped quote "" inside quoted field
          currentCell += '"';
          quotePendingClose = false;
          continue;
        }
        inQuotes = false;
        quotePendingClose = false;
        // Fall through to process delimiter / newline after closing quote
      }

      if (inQuotes) {
        if (ch === '"') {
          quotePendingClose = true;
        } else {
          currentCell += ch;
        }
        continue;
      }

      if (ch === '"' && currentCell.length === 0) {
        inQuotes = true;
        continue;
      }

      if (ch === ',') {
        currentRow.push(currentCell);
        currentCell = '';
        continue;
      }

      if (ch === '\r') {
        continue;
      }

      if (ch === '\n') {
        currentRow.push(currentCell);
        currentCell = '';
        // Skip completely blank trailing lines
        if (!(currentRow.length === 1 && currentRow[0] === '')) {
          yield currentRow;
        }
        currentRow = [];
        continue;
      }

      currentCell += ch;
    }
  }

  if (quotePendingClose) {
    inQuotes = false;
  }

  if (inQuotes) {
    throw new Error('Malformed CSV: unclosed quoted field.');
  }

  if (currentCell.length > 0 || currentRow.length > 0) {
    currentRow.push(currentCell);
    if (!(currentRow.length === 1 && currentRow[0] === '')) {
      yield currentRow;
    }
  }
}

async function convertJsonToCsv(inputPath: string, outputPath: string): Promise<void> {
  const ws = createWriteStream(outputPath, { encoding: 'utf8' });
  try {
    let headers: string[] | null = null;
    const headerSet = new Set<string>();

    for await (const item of streamJsonItems(inputPath)) {
      const record = toRecord(item);
      if (!headers) {
        headers = Object.keys(record);
        if (headers.length === 0) {
          headers = ['value'];
        }
        for (const h of headers) headerSet.add(h);
        await writeChunk(ws, headers.map(escapeCsvCell).join(',') + '\n');
      } else {
        // Capture values cleanly against established header columns
        for (const key of Object.keys(record)) {
          if (!headerSet.has(key)) {
            // Preserve unseen dynamic fields in an overflow column reference if needed,
            // or include matching keys. Keeping fixed CSV header guarantees column alignment.
          }
        }
      }

      const rowLine = headers.map((h) => escapeCsvCell(record[h])).join(',') + '\n';
      await writeChunk(ws, rowLine);
    }

    if (!headers) {
      // Valid empty JSON array `[]`
      await writeChunk(ws, '');
    }
    await closeWriteStream(ws);
  } catch (err) {
    ws.destroy();
    throw err;
  }
}

async function convertCsvToJson(inputPath: string, outputPath: string): Promise<void> {
  const ws = createWriteStream(outputPath, { encoding: 'utf8' });
  try {
    let headers: string[] | null = null;
    let rowCount = 0;

    await writeChunk(ws, '[');

    for await (const row of streamCsvRows(inputPath)) {
      if (!headers) {
        const seen = new Map<string, number>();
        headers = row.map((rawHeader, idx) => {
          const base = rawHeader.trim() || `column_${idx + 1}`;
          const count = seen.get(base) ?? 0;
          seen.set(base, count + 1);
          return count === 0 ? base : `${base}_${count + 1}`;
        });
        continue;
      }

      const obj: Record<string, string> = {};
      for (let i = 0; i < headers.length; i++) {
        obj[headers[i]] = row[i] ?? '';
      }
      // Include any extra columns beyond header width
      for (let i = headers.length; i < row.length; i++) {
        obj[`column_${i + 1}`] = row[i];
      }

      const prefix = rowCount === 0 ? '\n  ' : ',\n  ';
      await writeChunk(ws, prefix + JSON.stringify(obj));
      rowCount++;
    }

    await writeChunk(ws, rowCount > 0 ? '\n]\n' : ']\n');
    await closeWriteStream(ws);
  } catch (err) {
    ws.destroy();
    throw err;
  }
}

async function convertJsonToTxt(inputPath: string, outputPath: string): Promise<void> {
  const ws = createWriteStream(outputPath, { encoding: 'utf8' });
  try {
    let index = 0;
    for await (const item of streamJsonItems(inputPath)) {
      if (index > 0) {
        await writeChunk(ws, '\n---\n');
      }
      if (item !== null && typeof item === 'object' && !Array.isArray(item)) {
        const entries = Object.entries(item as Record<string, unknown>);
        const block = entries
          .map(([k, v]) => `${k}: ${typeof v === 'object' ? JSON.stringify(v) : String(v ?? '')}`)
          .join('\n');
        await writeChunk(ws, block + '\n');
      } else {
        await writeChunk(ws, JSON.stringify(item, null, 2) + '\n');
      }
      index++;
    }
    await closeWriteStream(ws);
  } catch (err) {
    ws.destroy();
    throw err;
  }
}

async function convertCsvToTxt(inputPath: string, outputPath: string): Promise<void> {
  const ws = createWriteStream(outputPath, { encoding: 'utf8' });
  try {
    for await (const row of streamCsvRows(inputPath)) {
      const sanitized = row.map((cell) => cell.replace(/[\t\r\n]+/g, ' '));
      await writeChunk(ws, sanitized.join('\t') + '\n');
    }
    await closeWriteStream(ws);
  } catch (err) {
    ws.destroy();
    throw err;
  }
}

async function convertTxtToJson(inputPath: string, outputPath: string): Promise<void> {
  const rs = createReadStream(inputPath, { encoding: 'utf8', highWaterMark: 64 * 1024 });
  const rl = readline.createInterface({ input: rs, crlfDelay: Infinity });
  const ws = createWriteStream(outputPath, { encoding: 'utf8' });

  try {
    let lineNumber = 0;
    await writeChunk(ws, '[');

    for await (let line of rl) {
      if (lineNumber === 0 && line.charCodeAt(0) === 0xfeff) {
        line = line.slice(1);
      }
      lineNumber++;
      const prefix = lineNumber === 1 ? '\n  ' : ',\n  ';
      await writeChunk(ws, prefix + JSON.stringify({ line: lineNumber, text: line }));
    }

    await writeChunk(ws, lineNumber > 0 ? '\n]\n' : ']\n');
    await closeWriteStream(ws);
  } catch (err) {
    rl.close();
    rs.destroy();
    ws.destroy();
    throw err;
  }
}

async function convertTxtToCsv(inputPath: string, outputPath: string): Promise<void> {
  const rs = createReadStream(inputPath, { encoding: 'utf8', highWaterMark: 64 * 1024 });
  const rl = readline.createInterface({ input: rs, crlfDelay: Infinity });
  const ws = createWriteStream(outputPath, { encoding: 'utf8' });

  try {
    let lineNumber = 0;
    await writeChunk(ws, 'line,text\n');

    for await (let line of rl) {
      if (lineNumber === 0 && line.charCodeAt(0) === 0xfeff) {
        line = line.slice(1);
      }
      lineNumber++;
      await writeChunk(ws, `${lineNumber},${escapeCsvCell(line)}\n`);
    }

    await closeWriteStream(ws);
  } catch (err) {
    rl.close();
    rs.destroy();
    ws.destroy();
    throw err;
  }
}

/**
 * Converts structured data & plain text (`json`, `csv`, `txt`) using native Node streams.
 */
export async function convertData(
  inputPath: string,
  outputPath: string,
  sourceFormat: string,
  targetFormat: string
): Promise<void> {
  if (!isDataFormat(sourceFormat) || !isDataFormat(targetFormat)) {
    throw new Error(`Unsupported data conversion route: ${sourceFormat} -> ${targetFormat}`);
  }
  if (sourceFormat === targetFormat) {
    throw new Error('Source and target formats must be different.');
  }

  const routeKey = `${sourceFormat}->${targetFormat}`;
  switch (routeKey) {
    case 'json->csv':
      return convertJsonToCsv(inputPath, outputPath);
    case 'csv->json':
      return convertCsvToJson(inputPath, outputPath);
    case 'json->txt':
      return convertJsonToTxt(inputPath, outputPath);
    case 'csv->txt':
      return convertCsvToTxt(inputPath, outputPath);
    case 'txt->json':
      return convertTxtToJson(inputPath, outputPath);
    case 'txt->csv':
      return convertTxtToCsv(inputPath, outputPath);
    default:
      throw new Error(`Unhandled data conversion route: ${sourceFormat} -> ${targetFormat}`);
  }
}
