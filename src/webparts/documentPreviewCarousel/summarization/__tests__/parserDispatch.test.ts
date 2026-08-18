import { isSummarizableExtension } from '../parsers/index';

describe('isSummarizableExtension', () => {
  it('accepts modern office formats', () => {
    expect(isSummarizableExtension('report.pdf')).toBe(true);
    expect(isSummarizableExtension('notes.docx')).toBe(true);
    expect(isSummarizableExtension('budget.xlsx')).toBe(true);
    expect(isSummarizableExtension('deck.pptx')).toBe(true);
  });

  it('accepts legacy extensions too (they get a clear error later, not a silent skip)', () => {
    expect(isSummarizableExtension('old.doc')).toBe(true);
    expect(isSummarizableExtension('old.xls')).toBe(true);
    expect(isSummarizableExtension('old.ppt')).toBe(true);
  });

  it('rejects non-document files', () => {
    expect(isSummarizableExtension('photo.png')).toBe(false);
    expect(isSummarizableExtension('archive.zip')).toBe(false);
    expect(isSummarizableExtension('noextension')).toBe(false);
  });

  it('is case-insensitive', () => {
    expect(isSummarizableExtension('REPORT.PDF')).toBe(true);
  });
});
