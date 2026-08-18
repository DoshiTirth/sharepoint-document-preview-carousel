/**
 * Excel (.xlsx) text extraction via read-excel-file.
 *
 * We flatten every sheet into a readable, row-by-row text representation
 * rather than trying to preserve exact tabular layout - the summarizer
 * downstream works on prose-like text, not spreadsheet grids. Sheet names
 * are kept as section headers so the model has structural context.
 *
 * Note on the API: as of read-excel-file v9, calling the default export
 * with no `sheet` option returns ALL sheets in one call (each with its name
 * and rows) rather than needing a separate "list sheet names" step.
 *
 * Legacy .xls (binary format) is NOT supported by read-excel-file - same
 * situation as legacy .doc. We surface that clearly rather than failing silently.
 */
import readXlsxFile, { Row, Sheet } from 'read-excel-file/browser';
import { DocumentParseError, IParsedDocument } from './types';

function rowToText(row: Row): string {
  return row
    .filter((cell) => cell !== null && cell !== undefined && String(cell).trim() !== '')
    .map((cell) => String(cell).trim())
    .join(' | ');
}

export async function parseXlsx(fileBytes: ArrayBuffer): Promise<IParsedDocument> {
  const blob = new Blob([fileBytes]);

  let sheets: Sheet[];
  try {
    sheets = await readXlsxFile(blob);
  } catch (error) {
    throw new DocumentParseError('Could not read this Excel file. It may be corrupted.', error);
  }

  const sections: string[] = [];

  for (const { sheet: sheetName, data: rows } of sheets) {
    const rowLines = rows.map(rowToText).filter((line) => line.length > 0);
    if (rowLines.length === 0) continue;
    sections.push(`## Sheet: ${sheetName}\n${rowLines.join('\n')}`);
  }

  const text = sections.join('\n\n');

  return {
    text,
    pageCount: sheets.length,
    isEffectivelyEmpty: text.trim().length < 20,
  };
}

export function isLegacyXlsFormat(fileName: string): boolean {
  return /\.xls$/i.test(fileName);
}
