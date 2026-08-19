# SharePoint Document Preview Carousel

A custom SharePoint Framework (SPFx) web part that turns a document library into a browsable, live-preview experience — instead of the default flat file list, viewers get a large preview window, clickable thumbnails, folder navigation with breadcrumbs, file-type filtering, and AI-powered document summarization, all in one component.

![version](https://img.shields.io/badge/version-2.0.0-green.svg)
![SPFx](https://img.shields.io/badge/SPFx-React-blue.svg)

## Why this exists

Standard SharePoint document libraries are functional but not very inviting — especially for content meant to be *browsed and presented*, like slide decks, reports, and dashboards, rather than just stored. This web part was built to make a document library feel more like a curated showcase: click a file, see it rendered immediately, without leaving the page. Version 2.0 adds a way to quickly understand a document's content without opening it in full, using an AI model that runs entirely on the viewer's own device.

## Features

- **Live inline preview** — click a file and it renders directly in a large preview pane. PowerPoint, Word, and Excel files render via SharePoint's built-in document viewer (WOPI); PDFs render natively in the browser.
- **AI document summarization** — a "Summarize AI" button generates a structured summary (overview + key points) of the currently selected PDF, Word, Excel, or PowerPoint document. Runs entirely in the viewer's browser using an in-browser language model (WebGPU) — no external AI service calls, no per-use cost, and no document content ever leaves the browser. Long documents are automatically broken into sections and summarized in stages; summaries are cached locally per file version so re-opening the same document is instant. See [How summarization works](#how-summarization-works) below for details, including hardware requirements and fallback behavior.
- **Folder navigation with breadcrumbs** — sub-folders show up as clickable items in the carousel itself; clicking one browses into it, with a breadcrumb trail to navigate back up. No fixed folder depth — works no matter how deeply nested a library's structure is.
- **Power BI report embedding** — pin one or more Power BI reports alongside regular files, configured directly in the web part's settings (no code changes needed). Access is enforced by each viewer's own Power BI permissions.
- **Live type filtering** — filter chips (PowerPoint / Word / Excel / PDF / Power BI / Other) let anyone viewing the page narrow down what's shown, on the fly. A separate setting controls which types are visible *by default* when the page loads.
- **Search** — filter the current folder's contents by file name.
- **Manual refresh** — re-check the current folder for new/removed files without reloading the page.
- **Fully permission-aware** — uses SharePoint's own authenticated REST API (`SPHttpClient`) under the current user's identity. It can never show a file or folder the viewer doesn't already have access to; permission errors surface as a clear message rather than a broken preview.
- **Handles real-world folder sizes** — automatically pages through SharePoint API results, so folders with hundreds of files still load completely rather than silently truncating.

## How summarization works

The "Summarize AI" button opens a panel above the preview with a structured summary of the currently selected document. A few things worth knowing about how it's built:

- **Runs entirely client-side.** The language model (a small, open-weight model — Llama 3.2 3B, quantized) downloads and runs directly in the browser via WebGPU. Nothing is sent to any external API or server — not the document content, not the summary. This also means there's no hosting cost and no usage limits to enforce.
- **Hardware requirement: WebGPU.** This needs a reasonably modern browser (Chrome or Edge, roughly 2023+). On a device or browser without WebGPU support, the panel shows a clear message explaining that instead of crashing.
- **First use downloads the model** (a few hundred MB to ~2GB), cached in the browser afterward — instant on repeat use, across all documents, not just the one you first opened it on.
- **Long documents are chunked.** The model's context window is limited (4096 tokens), so longer documents are automatically split into sections, each summarized individually, then combined into one final summary. The panel shows progress through this ("Summarizing section 3 of 12…") rather than a silent wait.
- **Summaries are cached per file version**, locally in your browser only (not shared with other viewers). Re-opening a summary for the same, unchanged document is instant; if the document is edited, the next summary request regenerates it.
- **Fallback behavior:** selecting a Power BI report, a folder, or nothing at all shows a clear explanatory message rather than an error. Documents with no extractable text (e.g. a scanned PDF with no text layer) are also reported clearly rather than producing an empty or broken summary.
- Supported formats: **PDF, Word (.docx), Excel (.xlsx), PowerPoint (.pptx)**. Legacy pre-2007 formats (.doc/.xls/.ppt) aren't supported — the panel will suggest re-saving in the modern format.

## How it's built

- **Framework:** SharePoint Framework (SPFx) 1.23, React, TypeScript
- **Build system:** Heft (Microsoft's newer SPFx build tooling, replacing gulp in recent SPFx versions)
- **Document preview:** SharePoint's native `WopiFrame.aspx` viewer for Office file types; direct browser rendering for PDFs — deliberately avoiding any public/external viewer service, since library content is typically private
- **Document summarization:** [WebLLM](https://github.com/mlc-ai/web-llm) for in-browser model inference; `pdf.js` for PDF text extraction; `mammoth` for Word; `read-excel-file` for Excel; a small custom OOXML parser (via JSZip) for PowerPoint, since no lightweight library exists for that format
- **Data access:** `SPHttpClient` calling SharePoint's REST API (`_api/web/GetFolderByServerRelativeUrl`), with pagination handling for large folders and OData-safe escaping for folder/file names containing apostrophes

## Configuration

All configuration happens through the web part's property pane in SharePoint — no code editing required to use it:

| Setting | What it does |
|---|---|
| Folder server-relative URL | The folder the carousel starts browsing from, e.g. `/sites/YourSiteName/Shared Documents/Your Folder` |
| Power BI reports | One report per line: `Report Name \| Embed URL` (use Power BI's "Embed report → SharePoint Online" link) |
| Default filters | Checkboxes controlling which file types are visible when the page first loads (viewers can still adjust live) |

## Getting started (development)

```bash
npm install
npm run build
```

This produces a `.sppkg` package under `sharepoint/solution/`, ready to upload to a SharePoint App Catalog and deploy to any site in a tenant.

> Built and tested using GitHub Codespaces (Node 22 LTS via nvm) as the development environment.

## Disclaimer

This code is provided as-is, without warranty of any kind. Built as a personal project to explore SPFx development patterns for document-centric SharePoint experiences.
