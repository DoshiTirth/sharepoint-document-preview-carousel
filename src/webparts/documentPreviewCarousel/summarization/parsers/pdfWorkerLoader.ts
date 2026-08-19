/**
 * Loads the PDF.js worker script for use in an SPFx web part.
 *
 * WHY THIS EXISTS (read before touching this file):
 * SPFx's webpack/Heft toolchain is built around CommonJS-era conventions and
 * has repeatedly broken on pdf.js's modern ES-module worker builds (see
 * SharePoint/sp-dev-docs#8896 and #9727 - both show webpack choking trying
 * to parse pdf.worker.mjs directly). Importing the worker normally
 * (`import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs'`) hands that file
 * to webpack's JS pipeline, which is exactly what breaks.
 *
 * Workaround: the worker's source is committed as a plain .txt asset
 * (see assets/pdf-worker/pdf-worker-source.txt) so webpack's static-asset
 * loader just copies it byte-for-byte without trying to parse it as a
 * module. At runtime we fetch that file's *deployed* URL, read it as text,
 * and turn it into a Blob URL - which is a completely valid worker script
 * source as far as the browser's Worker API is concerned. This is bundler-
 * agnostic and doesn't depend on SPFx ever correctly understanding the
 * worker's ES module syntax.
 *
 * If you upgrade pdfjs-dist, you MUST re-copy the new legacy worker build:
 *   cp node_modules/pdfjs-dist/legacy/build/pdf.worker.min.mjs \
 *      src/webparts/documentPreviewCarousel/assets/pdf-worker/pdf-worker-source.txt
 */

// SPFx's webpack config resolves import of asset files to their deployed
// URL string (same mechanism used for the web part's icon/images elsewhere
// in this project). See ../assets.d.ts for the module declaration this relies on.
import workerSourceAssetUrl from '../../assets/pdf-worker/pdf-worker-source.txt';

let cachedBlobUrl: string | undefined;
let inFlightRequest: Promise<string> | undefined;

/**
 * Returns a Blob URL that can be assigned to pdfjsLib.GlobalWorkerOptions.workerSrc.
 * Cached after the first successful call - the fetch + Blob creation only
 * happens once per browser session.
 */
export async function getPdfWorkerBlobUrl(): Promise<string> {
  if (cachedBlobUrl) return cachedBlobUrl;
  if (inFlightRequest) return inFlightRequest;

  inFlightRequest = (async () => {
    const response = await fetch(workerSourceAssetUrl);
    if (!response.ok) {
      throw new Error(`Failed to fetch PDF worker asset (status ${response.status}).`);
    }
    const workerSourceText = await response.text();
    const blob = new Blob([workerSourceText], { type: 'application/javascript' });
    cachedBlobUrl = URL.createObjectURL(blob);
    return cachedBlobUrl;
  })();

  return inFlightRequest;
}

/**
 * Releases the cached Blob URL's underlying memory (a fixed ~1.4MB - the
 * worker script's text). Safe to call even if the URL was never created
 * (checks internally, does nothing in that case). Intended to be called
 * from the web part's onDispose(), same as the model engine's unload -
 * without this, the Blob's content would stay retained in memory for the
 * lifetime of the page even after this web part instance is gone.
 */
export function revokePdfWorkerBlobUrl(): void {
  if (cachedBlobUrl) {
    URL.revokeObjectURL(cachedBlobUrl);
    cachedBlobUrl = undefined;
  }
}
