/**
 * read-excel-file is mocked here. It reads Excel files via the browser
 * Blob API (calling blob.arrayBuffer() internally), and this project's
 * jsdom test environment's Blob implementation doesn't provide that method
 * at all (confirmed directly: `typeof new Blob([...]).arrayBuffer` is
 * `undefined` here) - every real browser has supported it for years, so
 * this is a gap in this jsdom version, not a bug in our code or in
 * read-excel-file.
 *
 * What we actually own and need to verify is OUR row-flattening and
 * section-building logic (joining cells with " | ", skipping blank rows,
 * grouping by sheet name) - not read-excel-file's own XLSX-parsing
 * correctness, which is its own test suite's responsibility.
 */

jest.mock('read-excel-file/browser', () => jest.fn());

import readXlsxFile from 'read-excel-file/browser';
import { parseXlsx, isLegacyXlsFormat } from '../parsers/xlsxParser';
import { DocumentParseError } from '../parsers/types';

describe('parseXlsx', () => {
  it('flattens rows from multiple sheets into labeled sections', async () => {
    (readXlsxFile as jest.Mock).mockResolvedValue([
      {
        sheet: 'Revenue',
        data: [
          ['Region', 'Q1', 'Q2'],
          ['North', 1200, 1350],
        ],
      },
      {
        sheet: 'Notes',
        data: [['Follow up with South region team about slower growth.']],
      },
    ]);

    const result = await parseXlsx(new ArrayBuffer(0));

    expect(result.isEffectivelyEmpty).toBe(false);
    expect(result.pageCount).toBe(2);
    expect(result.text).toContain('## Sheet: Revenue');
    expect(result.text).toContain('Region | Q1 | Q2');
    expect(result.text).toContain('North | 1200 | 1350');
    expect(result.text).toContain('## Sheet: Notes');
    expect(result.text).toContain('Follow up with South region team');
  });

  it('skips sheets that are entirely blank rather than emitting an empty section', async () => {
    (readXlsxFile as jest.Mock).mockResolvedValue([
      { sheet: 'HasData', data: [['x']] },
      { sheet: 'Empty', data: [[], [null, undefined, '']] },
    ]);

    const result = await parseXlsx(new ArrayBuffer(0));

    expect(result.text).toContain('## Sheet: HasData');
    expect(result.text).not.toContain('## Sheet: Empty');
  });

  it('reports isEffectivelyEmpty when every sheet is blank', async () => {
    (readXlsxFile as jest.Mock).mockResolvedValue([{ sheet: 'Empty', data: [[]] }]);

    const result = await parseXlsx(new ArrayBuffer(0));

    expect(result.isEffectivelyEmpty).toBe(true);
  });

  it('throws DocumentParseError when the file fails to load', async () => {
    (readXlsxFile as jest.Mock).mockRejectedValue(new Error('corrupt file'));

    await expect(parseXlsx(new ArrayBuffer(0))).rejects.toThrow(DocumentParseError);
  });
});

describe('isLegacyXlsFormat', () => {
  it('detects legacy .xls extension', () => {
    expect(isLegacyXlsFormat('budget.xls')).toBe(true);
    expect(isLegacyXlsFormat('budget.xlsx')).toBe(false);
  });
});
