/**
 * Single entry point for document text extraction. Routes to the correct
 * format-specific parser based on file extension, and normalizes legacy
 * (pre-2007 binary) formats into a clear, actionable error rather than a
 * confusing crash deep in a parsing library.
 */
import { parseDocx, isLegacyDocFormat } from './docxParser';
import { parsePptx, isLegacyPptFormat } from './pptxParser';
import { DocumentParseError, IParsedDocument } from './types';
import { isLegacyXlsFormat, parseXlsx } from './xlsxParser';

export { DocumentParseError, IParsedDocument } from './types';

export interface IParseFileOptions {
  /** Advanced/testing use only - overrides the default bundled worker loader. */
  pdfWorkerSrc?: string;
}

function getExtension(fileName: string): string {
  const parts = fileName.split('.');
  return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : '';
}

export function isSummarizableExtension(fileName: string): boolean {
  return ['pdf', 'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt'].indexOf(getExtension(fileName)) !== -1;
}

/**
 * Parse a document's raw bytes into plain text.
 *
 * Throws DocumentParseError for anything that prevents extraction:
 * unsupported legacy formats, corrupted files, or files with no text.
 * Callers (the UI layer) are expected to catch this and show a fallback
 * message rather than let it propagate as an unhandled crash.
 */
export async function parseFile(
  fileName: string,
  fileBytes: ArrayBuffer,
  options: IParseFileOptions = {}
): Promise<IParsedDocument> {
  const extension = getExtension(fileName);

  if (isLegacyDocFormat(fileName) || isLegacyPptFormat(fileName) || isLegacyXlsFormat(fileName)) {
    throw new DocumentParseError(
      `Legacy ".${extension}" files (pre-2007 format) aren't supported for summarization. ` +
        'Try re-saving this file in the modern format (.docx/.xlsx/.pptx) and try again.'
    );
  }

  switch (extension) {
    case 'pdf': {
      // Lazy-loaded: pdfjs-dist is a large, ESM-only dependency. Loading it
      // only when a PDF is actually opened keeps it out of the bundle's
      // critical path for people summarizing Word/Excel/PowerPoint files,
      // and keeps it out of the module graph for anything (like tests) that
      // only needs the lighter-weight parsers.
      const [{ parsePdf }, { getPdfWorkerBlobUrl }] = await Promise.all([
        import(/* webpackChunkName: 'pdf-parser' */ './pdfParser'),
        import(/* webpackChunkName: 'pdf-parser' */ './pdfWorkerLoader'),
      ]);
      let workerSrc: string;
      try {
        workerSrc = options.pdfWorkerSrc || (await getPdfWorkerBlobUrl());
      } catch (error) {
        throw new DocumentParseError('Could not initialize the PDF reader. Please try again.', error);
      }
      return parsePdf(fileBytes, { workerSrc });
    }
    case 'docx':
      return parseDocx(fileBytes);
    case 'xlsx':
      return parseXlsx(fileBytes);
    case 'pptx':
      return parsePptx(fileBytes);
    default:
      throw new DocumentParseError(`Unsupported file type ".${extension}" for summarization.`);
  }
}
