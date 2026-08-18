/**
 * These tests mock mammoth rather than exercising it against a real .docx
 * fixture under Jest. mammoth ships genuinely different code paths for
 * Node (`{ buffer }`) vs browser bundlers (`{ arrayBuffer }`), selected via
 * package.json's "browser" field. Real webpack bundling (confirmed by
 * reading mammoth's browser/unzip.js source directly) resolves to the
 * `{ arrayBuffer }` variant, which is what parseDocx correctly calls - but
 * Jest's module resolution doesn't apply that same "browser" field
 * remapping, so a real fixture would exercise the wrong code path here and
 * fail for reasons that don't reflect an actual production bug.
 *
 * What we actually own and need to verify is OUR integration logic: that
 * we call mammoth correctly, normalize its output, and wrap failures as
 * DocumentParseError. Whether mammoth itself correctly parses OOXML is
 * mammoth's own test suite's job.
 */

jest.mock('mammoth', () => ({
  extractRawText: jest.fn(),
}));

import * as mammoth from 'mammoth';
import { parseDocx, isLegacyDocFormat } from '../parsers/docxParser';
import { DocumentParseError } from '../parsers/types';

describe('parseDocx', () => {
  it('normalizes mammoth output and collapses excess blank lines', async () => {
    (mammoth.extractRawText as jest.Mock).mockResolvedValue({
      value: 'Quarterly Report\n\n\n\nRevenue grew fourteen percent.\n',
      messages: [],
    });

    const result = await parseDocx(new ArrayBuffer(0));

    expect(result.isEffectivelyEmpty).toBe(false);
    expect(result.text).toContain('Quarterly Report');
    expect(result.text).toContain('Revenue grew fourteen percent.');
    expect(result.text).not.toContain('\n\n\n');
  });

  it('flags near-empty extraction as isEffectivelyEmpty', async () => {
    (mammoth.extractRawText as jest.Mock).mockResolvedValue({ value: '  ', messages: [] });

    const result = await parseDocx(new ArrayBuffer(0));

    expect(result.isEffectivelyEmpty).toBe(true);
  });

  it('wraps a mammoth failure as DocumentParseError', async () => {
    (mammoth.extractRawText as jest.Mock).mockRejectedValue(new Error('corrupt zip'));

    await expect(parseDocx(new ArrayBuffer(0))).rejects.toThrow(DocumentParseError);
  });

  it('calls mammoth with the browser-style arrayBuffer input shape', async () => {
    (mammoth.extractRawText as jest.Mock).mockResolvedValue({ value: 'text', messages: [] });
    const bytes = new ArrayBuffer(4);

    await parseDocx(bytes);

    expect(mammoth.extractRawText).toHaveBeenCalledWith({ arrayBuffer: bytes });
  });
});

describe('isLegacyDocFormat', () => {
  it('detects legacy .doc extension', () => {
    expect(isLegacyDocFormat('report.doc')).toBe(true);
    expect(isLegacyDocFormat('report.docx')).toBe(false);
  });
});
