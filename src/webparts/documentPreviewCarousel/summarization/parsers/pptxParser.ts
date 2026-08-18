/**
 * PowerPoint (.pptx) text extraction.
 *
 * .pptx is a zip archive containing one XML file per slide under
 * ppt/slides/slideN.xml. There's no lightweight npm library that does this
 * well for browser use, so we unzip with JSZip (already a transitive
 * dependency via mammoth/SPFx tooling) and pull text out of <a:t> runs
 * ourselves using the browser's built-in DOMParser - no extra XML library needed.
 *
 * Legacy .ppt (binary format) is NOT supported - same situation as .doc/.xls.
 */
import JSZip from 'jszip';
import { DocumentParseError, IParsedDocument } from './types';

function extractTextFromSlideXml(xml: string): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');

  const parserError = doc.querySelector('parsererror');
  if (parserError) return '';

  // Text runs in OOXML slides live in <a:t> elements, regardless of what
  // shape/placeholder they're nested inside (title, body, table cell, etc).
  const textNodes = Array.from(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't'));

  return textNodes
    .map((node) => node.textContent || '')
    .filter((t) => t.trim().length > 0)
    .join(' ');
}

function getSlideNumberFromPath(path: string): number {
  const match = /slide(\d+)\.xml$/.exec(path);
  return match ? parseInt(match[1], 10) : Number.MAX_SAFE_INTEGER;
}

export async function parsePptx(fileBytes: ArrayBuffer): Promise<IParsedDocument> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(fileBytes);
  } catch (error) {
    throw new DocumentParseError('Could not read this PowerPoint file. It may be corrupted.', error);
  }

  const slidePaths = Object.keys(zip.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
    .sort((a, b) => getSlideNumberFromPath(a) - getSlideNumberFromPath(b));

  if (slidePaths.length === 0) {
    throw new DocumentParseError('This PowerPoint file has no readable slides.');
  }

  const slideTexts: string[] = [];

  for (let i = 0; i < slidePaths.length; i++) {
    const path = slidePaths[i];
    try {
      const xml = await zip.files[path].async('text');
      const slideText = extractTextFromSlideXml(xml);
      if (slideText.trim().length > 0) {
        slideTexts.push(`## Slide ${i + 1}\n${slideText.trim()}`);
      }
    } catch (error) {
      console.warn(`[ctxpack-summarizer] Skipping unreadable slide "${path}"`, error);
    }
  }

  const text = slideTexts.join('\n\n');

  return {
    text,
    pageCount: slidePaths.length,
    isEffectivelyEmpty: text.trim().length < 20,
  };
}

export function isLegacyPptFormat(fileName: string): boolean {
  return /\.ppt$/i.test(fileName);
}
