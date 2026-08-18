/**
 * PDF text extraction via PDF.js.
 *
 * PDF.js needs its "worker" script available to do parsing off the main
 * thread. In an SPFx web part we bundle the worker as a static asset and
 * point PDF.js at it explicitly (rather than relying on a CDN, which would
 * be a network dependency and a CSP headache in SharePoint).
 */
import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
import { DocumentParseError, IParsedDocument } from './types';

let workerConfigured = false;

function ensureWorkerConfigured(workerSrc: string): void {
  if (workerConfigured) return;
  pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;
  workerConfigured = true;
}

export interface IPdfParseOptions {
  /** URL to the pdf.js worker script, typically bundled under the web part's assets. */
  workerSrc: string;
}

export async function parsePdf(fileBytes: ArrayBuffer, options: IPdfParseOptions): Promise<IParsedDocument> {
  ensureWorkerConfigured(options.workerSrc);

  let doc: PDFDocumentProxy;
  try {
    doc = await pdfjsLib.getDocument({ data: fileBytes }).promise;
  } catch (error) {
    throw new DocumentParseError('Could not read this PDF. It may be corrupted or password-protected.', error);
  }

  const pageTexts: string[] = [];
  try {
    for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
      const page = await doc.getPage(pageNum);
      const content = await page.getTextContent();
      const pageText = content.items
        .map((item) => ('str' in item ? item.str : ''))
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
      pageTexts.push(pageText);
    }
  } catch (error) {
    throw new DocumentParseError('Failed while extracting text from the PDF.', error);
  } finally {
    doc.destroy().catch(() => undefined);
  }

  const text = pageTexts.join('\n\n');
  const isEffectivelyEmpty = text.trim().length < 20;

  return {
    text,
    pageCount: doc.numPages,
    isEffectivelyEmpty,
  };
}
