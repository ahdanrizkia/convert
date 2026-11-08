declare module 'mammoth' {
  export interface MammothResult {
    value: string;
    messages: unknown[];
  }
  export function extractRawText(options: { path?: string; buffer?: Buffer }): Promise<MammothResult>;
  export function convertToHtml(options: { path?: string; buffer?: Buffer }): Promise<MammothResult>;
}

declare module 'pdf-parse' {
  export interface PDFParseResult {
    pages: Array<{ text: string; num: number }>;
    text: string;
    total: number;
  }
  export interface ScreenshotPage {
    data: Uint8Array;
    dataUrl: string;
    pageNumber: number;
    width: number;
    height: number;
    scale: number;
  }
  export interface ScreenshotResult {
    pages: ScreenshotPage[];
    total: number;
  }
  export class PDFParse {
    constructor(options: { data: Buffer | Uint8Array; verbosity?: number });
    getText(): Promise<PDFParseResult>;
    getScreenshot(options?: { pageNumber?: number; scale?: number }): Promise<ScreenshotResult>;
    destroy(): Promise<void>;
  }
}

declare module 'pdfjs-dist/legacy/build/pdf.mjs' {
  export const GlobalWorkerOptions: { workerSrc: string };
  export function getDocument(src: unknown): {
    promise: Promise<{
      numPages: number;
      getPage(pageNumber: number): Promise<{
        getViewport(params: { scale: number }): { width: number; height: number };
        render(params: { canvasContext: unknown; viewport: unknown }): { promise: Promise<void> };
      }>;
    }>;
  };
}
