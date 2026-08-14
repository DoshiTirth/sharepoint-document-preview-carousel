# SharePoint Document Preview Carousel

A custom SharePoint Framework (SPFx) web part that turns a document library into a browsable, live-preview experience — instead of the default flat file list, viewers get a large preview window, clickable thumbnails, folder navigation with breadcrumbs, and file-type filtering, all in one component.

![version](https://img.shields.io/badge/version-1.23.2-green.svg)
![SPFx](https://img.shields.io/badge/SPFx-React-blue.svg)

## Why this exists

Standard SharePoint document libraries are functional but not very inviting — especially for content meant to be *browsed and presented*, like slide decks, reports, and dashboards, rather than just stored. This web part was built to make a document library feel more like a curated showcase: click a file, see it rendered immediately, without leaving the page.

## Features

- **Live inline preview** — click a file and it renders directly in a large preview pane. PowerPoint, Word, and Excel files render via SharePoint's built-in document viewer (WOPI); PDFs render natively in the browser.
- **Folder navigation with breadcrumbs** — sub-folders show up as clickable items in the carousel itself; clicking one browses into it, with a breadcrumb trail to navigate back up. No fixed folder depth — works no matter how deeply nested a library's structure is.
- **Power BI report embedding** — pin one or more Power BI reports alongside regular files, configured directly in the web part's settings (no code changes needed). Access is enforced by each viewer's own Power BI permissions.
- **Live type filtering** — filter chips (PowerPoint / Word / Excel / PDF / Power BI / Other) let anyone viewing the page narrow down what's shown, on the fly. A separate setting controls which types are visible *by default* when the page loads.
- **Search** — filter the current folder's contents by file name.
- **Manual refresh** — re-check the current folder for new/removed files without reloading the
cat README.md
cat > README.md << 'EOF'
# SharePoint Document Preview Carousel

A custom SharePoint Framework (SPFx) web part that turns a document library into a browsable, live-preview experience — instead of the default flat file list, viewers get a large preview window, clickable thumbnails, folder navigation with breadcrumbs, and file-type filtering, all in one component.

![version](https://img.shields.io/badge/version-1.23.2-green.svg)
![SPFx](https://img.shields.io/badge/SPFx-React-blue.svg)

## Why this exists

Standard SharePoint document libraries are functional but not very inviting — especially for content meant to be *browsed and presented*, like slide decks, reports, and dashboards, rather than just stored. This web part was built to make a document library feel more like a curated showcase: click a file, see it rendered immediately, without leaving the page.

## Features

- **Live inline preview** — click a file and it renders directly in a large preview pane. PowerPoint, Word, and Excel files render via SharePoint's built-in document viewer (WOPI); PDFs render natively in the browser.
- **Folder navigation with breadcrumbs** — sub-folders show up as clickable items in the carousel itself; clicking one browses into it, with a breadcrumb trail to navigate back up. No fixed folder depth — works no matter how deeply nested a library's structure is.
- **Power BI report embedding** — pin one or more Power BI reports alongside regular files, configured directly in the web part's settings (no code changes needed). Access is enforced by each viewer's own Power BI permissions.
- **Live type filtering** — filter chips (PowerPoint / Word / Excel / PDF / Power BI / Other) let anyone viewing the page narrow down what's shown, on the fly. A separate setting controls which types are visible *by default* when the page loads.
- **Search** — filter the current folder's contents by file name.
- **Manual refresh** — re-check the current folder for new/removed files without reloading the page.
- **Fully permission-aware** — uses SharePoint's own authenticated REST API (`SPHttpClient`) under the current user's identity. It can never show a file or folder the viewer doesn't already have access to; permission errors surface as a clear message rather than a broken preview.
- **Handles real-world folder sizes** — automatically pages through SharePoint API results, so folders with hundreds of files still load completely rather than silently truncating.

## How it's built

- **Framework:** SharePoint Framework (SPFx) 1.23, React, TypeScript
- **Build system:** Heft (Microsoft's newer SPFx build tooling, replacing gulp in recent SPFx versions)
- **Document preview:** SharePoint's native `WopiFrame.aspx` viewer for Office file types; direct browser rendering for PDFs — deliberately avoiding any public/external viewer service, since library content is typically private
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
