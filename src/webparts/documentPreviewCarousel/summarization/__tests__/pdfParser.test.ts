/**
 * These tests mock pdfjs-dist entirely rather than exercising real PDF
 * parsing. Two reasons:
 *
 * 1. pdfjs-dist ships ESM-only (no CommonJS build) as of the version this
 *    project uses, which is a well-known incompatibility with Jest's default
 *    CommonJS-style module execution - not something specific to this
 *    project's setup. Getting real parsing to run under Jest would require
 *    fighting Jest's transform pipeline for a well-tested third-party
 *    library, for little benefit.
 * 2. What we actually own and need to verify is OUR integration logic:
 *    looping over pages in order, joining their text, computing
 *    isEffectivelyEmpty, always calling doc.destroy() even on error, and
 *    wrapping failures as DocumentParseError. Whether pdf.js itself
 *    correctly parses PDF byte structure is pdf.js's own test suite's job,
 *    not ours to re-verify.
 *
 * Real end-to-end PDF parsing (including the worker-loading path) should
 * still be manually verified against the SharePoint Workbench / a real
 * tenant before this ships, since that's the one piece of this module that
 * can't be meaningfully unit-tested in this environment.
 */

const mockDestroy = jest.fn().mockResolvedValue(undefined);

function makeMockPage(text: string): { getTextContent: () => Promise<{ items: { str: string }[] }> } {
  return {
    getTextContent: () => Promise.resolve({ items: [{ str: text }] }),
  };
}

jest.mock('pdfjs-dist', () => ({
  GlobalWorkerOptions: {},
  getDocument: jest.fn(),
}));

import * as pdfjsLib from 'pdfjs-dist';
import { parsePdf } from '../parsers/pdfParser';
import { DocumentParseError } from '../parsers/types';

describe('parsePdf', () => {
  beforeEach(() => {
    mockDestroy.mockClear();
  });

  it('joins text from all pages in order and reports page count', async () => {
    const mockDoc = {
      numPages: 2,
      getPage: jest.fn((pageNum: number) =>
        Promise.resolve(makeMockPage(pageNum === 1 ? 'First page text' : 'Second page text'))
      ),
      destroy: mockDestroy,
    };
    (pdfjsLib.getDocument as jest.Mock).mockReturnValue({ promise: Promise.resolve(mockDoc) });

    const result = await parsePdf(new ArrayBuffer(0), { workerSrc: 'irrelevant-for-this-test' });

    expect(result.pageCount).toBe(2);
    expect(result.text).toContain('First page text');
    expect(result.text).toContain('Second page text');
    expect(result.isEffectivelyEmpty).toBe(false);
    expect(mockDestroy).toHaveBeenCalledTimes(1);
  });

  it('flags near-empty extraction as isEffectivelyEmpty (e.g. a scanned PDF with no text layer)', async () => {
    const mockDoc = {
      numPages: 1,
      getPage: jest.fn(() => Promise.resolve(makeMockPage(''))),
      destroy: mockDestroy,
    };
    (pdfjsLib.getDocument as jest.Mock).mockReturnValue({ promise: Promise.resolve(mockDoc) });

    const result = await parsePdf(new ArrayBuffer(0), { workerSrc: 'irrelevant-for-this-test' });

    expect(result.isEffectivelyEmpty).toBe(true);
  });

  it('wraps a document-load failure as DocumentParseError', async () => {
    (pdfjsLib.getDocument as jest.Mock).mockReturnValue({
      promise: Promise.reject(new Error('corrupt bytes')),
    });

    await expect(parsePdf(new ArrayBuffer(0), { workerSrc: 'irrelevant' })).rejects.toThrow(
      DocumentParseError
    );
  });

  it('still calls destroy() even when text extraction fails partway through', async () => {
    const mockDoc = {
      numPages: 1,
      getPage: jest.fn(() => Promise.reject(new Error('extraction failed'))),
      destroy: mockDestroy,
    };
    (pdfjsLib.getDocument as jest.Mock).mockReturnValue({ promise: Promise.resolve(mockDoc) });

    await expect(parsePdf(new ArrayBuffer(0), { workerSrc: 'irrelevant' })).rejects.toThrow(
      DocumentParseError
    );
    expect(mockDestroy).toHaveBeenCalledTimes(1);
  });
});
