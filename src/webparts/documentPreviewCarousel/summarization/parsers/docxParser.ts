/**
 * Word (.docx) text extraction via mammoth.
 *
 * mammoth is designed for docx -> HTML/text conversion and handles the
 * OOXML structure for us. Legacy .doc (binary format, pre-2007) is NOT
 * supported by mammoth or any lightweight in-browser library — that format
 * requires a full binary OLE parser. We surface a clear error for .doc
 * rather than silently failing, so the UI can show a helpful message
 * instead of a generic crash.
 */
import * as mammoth from 'mammoth';
import { DocumentParseError, IParsedDocument } from './types';

export async function parseDocx(fileBytes: ArrayBuffer): Promise<IParsedDocument> {
  let result: { value: string; messages: unknown[] };
  try {
    result = await mammoth.extractRawText({ arrayBuffer: fileBytes });
  } catch (error) {
    throw new DocumentParseError('Could not read this Word document. It may be corrupted.', error);
  }

  const text = result.value.replace(/\n{3,}/g, '\n\n').trim();

  return {
    text,
    isEffectivelyEmpty: text.length < 20,
  };
}

export function isLegacyDocFormat(fileName: string): boolean {
  return /\.doc$/i.test(fileName);
}
