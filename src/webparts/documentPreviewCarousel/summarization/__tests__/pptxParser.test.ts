/**
 * The zip-loading step (JSZip.loadAsync) is mocked here rather than run
 * against a real .pptx fixture. Confirmed via a standalone script that the
 * exact same fixture + conversion logic loads correctly through JSZip in
 * plain Node - the failure only appears under Jest's jsdom test environment,
 * which runs in a separate JS "realm" with its own ArrayBuffer/Uint8Array
 * globals. JSZip's internal type-checking doesn't recognize objects from a
 * different realm as valid, even though they're structurally identical.
 * This is a known jsdom/Jest cross-realm quirk, not a bug in our code or in
 * JSZip - confirmed by the fact that it works in a real browser and in
 * plain Node.
 *
 * What we mock is only the zip-file-loading step. The actual OOXML slide
 * text extraction (extractTextFromSlideXml, using DOMParser) still runs for
 * real against realistic slide XML below - DOMParser is something jsdom
 * genuinely implements correctly, unlike some newer Blob/File APIs.
 */

function slideXml(texts: string[]): string {
  const runs = texts
    .map(
      (t) =>
        `<a:p><a:r><a:t>${t}</a:t></a:r></a:p>`
    )
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody>${runs}</p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>`;
}

const mockFiles: Record<string, { async: (type: string) => Promise<string> }> = {};

jest.mock('jszip', () => ({
  loadAsync: jest.fn(() => Promise.resolve({ files: mockFiles })),
}));

import JSZip from 'jszip';
import { parsePptx, isLegacyPptFormat } from '../parsers/pptxParser';
import { DocumentParseError } from '../parsers/types';

function setMockSlides(slideTextsByNumber: Record<number, string[]>): void {
  for (const key of Object.keys(mockFiles)) delete mockFiles[key];
  for (const [num, texts] of Object.entries(slideTextsByNumber)) {
    mockFiles[`ppt/slides/slide${num}.xml`] = {
      async: () => Promise.resolve(slideXml(texts)),
    };
  }
}

describe('parsePptx', () => {
  it('extracts text from all slides in numeric order, regardless of file listing order', async () => {
    setMockSlides({
      2: ['Key Priorities', 'Improve onboarding time'],
      1: ['Annual Kickoff', 'Team goals for the year ahead'],
    });

    const result = await parsePptx(new ArrayBuffer(0));

    expect(result.isEffectivelyEmpty).toBe(false);
    expect(result.pageCount).toBe(2);
    // Slide 1's content must appear before slide 2's, proving numeric sort
    // (not file-listing order, which was reversed above) drives the output.
    const slide1Index = result.text.indexOf('Annual Kickoff');
    const slide2Index = result.text.indexOf('Key Priorities');
    expect(slide1Index).toBeGreaterThanOrEqual(0);
    expect(slide2Index).toBeGreaterThan(slide1Index);
  });

  it('ignores non-slide files in the zip (e.g. slide layouts, media)', async () => {
    setMockSlides({ 1: ['Only slide'] });
    mockFiles['ppt/slideLayouts/slideLayout1.xml'] = {
      async: () => Promise.resolve(slideXml(['Should not appear'])),
    };

    const result = await parsePptx(new ArrayBuffer(0));

    expect(result.text).toContain('Only slide');
    expect(result.text).not.toContain('Should not appear');
    expect(result.pageCount).toBe(1);
  });

  it('throws DocumentParseError when there are no slides at all', async () => {
    setMockSlides({});

    await expect(parsePptx(new ArrayBuffer(0))).rejects.toThrow(
      'This PowerPoint file has no readable slides'
    );
  });

  it('throws DocumentParseError when the zip itself fails to load', async () => {
    (JSZip.loadAsync as jest.Mock).mockRejectedValueOnce(new Error('not a zip'));

    await expect(parsePptx(new ArrayBuffer(0))).rejects.toThrow(DocumentParseError);
  });
});

describe('isLegacyPptFormat', () => {
  it('detects legacy .ppt extension', () => {
    expect(isLegacyPptFormat('deck.ppt')).toBe(true);
    expect(isLegacyPptFormat('deck.pptx')).toBe(false);
  });
});
