/**
 * Shared types for document text extraction.
 *
 * All parsers (PDF, Word, Excel, PowerPoint) normalize to this same shape so
 * the summarization pipeline downstream doesn't need to know which file
 * format it started from.
 */

export interface IParsedDocument {
  /** Plain text content, in reading order as best as the format allows. */
  text: string;
  /** Number of "pages" or "sections" the source had, if the format has that concept. */
  pageCount?: number;
  /** True if extraction succeeded but the document appears to have little/no text
   *  (e.g. a scanned PDF with no text layer, or an empty spreadsheet). */
  isEffectivelyEmpty: boolean;
}

export type SupportedParserExtension = 'pdf' | 'docx' | 'doc' | 'xlsx' | 'xls' | 'pptx' | 'ppt';

export class DocumentParseError extends Error {
  constructor(message: string, public readonly cause?: unknown) {
    super(message);
    this.name = 'DocumentParseError';
    // TypeScript compiles this project to ES5 (see tsconfig), and ES5's
    // downleveled `class X extends Error` transpilation does not correctly
    // preserve the prototype chain - without this line, `instanceof
    // DocumentParseError` returns false even for errors constructed by this
    // exact class. This is a well-documented TypeScript/ES5 gotcha for any
    // class extending a native built-in (Error, Array, Map, etc).
    Object.setPrototypeOf(this, DocumentParseError.prototype);
  }
}
