import * as React from 'react';
import { IDocumentPreviewCarouselProps } from './IDocumentPreviewCarouselProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { ISummaryTarget } from '../summarization/ui/SummaryPanel';

// Lazy-loaded: the summarization feature (and everything it pulls in -
// mammoth, JSZip, read-excel-file, and eventually pdf.js/WebLLM further
// down its own lazy chain) should not cost anything on initial page load
// for people who never click "Summarize AI". Only SummaryPanel's *type*
// (imported above, erased at compile time) is needed eagerly for props typing.
const SummaryPanel = React.lazy(() =>
  import(/* webpackChunkName: 'summary-panel' */ '../summarization/ui/SummaryPanel').then((mod) => ({
    default: mod.SummaryPanel,
  }))
);
/* ============================================================
   WHAT THIS COMPONENT DOES (read this first)
   ============================================================
   This shows a big "preview" box at the top, and a row of
   clickable file/folder thumbnails ("the carousel") underneath.

   - Click a FILE thumbnail   -> it loads into the big preview box.
   - Click the big preview    -> opens that file in a new tab.
   - Click a FOLDER thumbnail -> the carousel goes INSIDE that
     folder and shows what's in there instead.
   - Breadcrumbs above the carousel show where you are, and let
     you click back up to a parent folder.
   - A search box lets you filter what's shown by name.
   - A Refresh button re-checks the folder for new/removed files
     without needing to reload the whole page.

   Nothing is auto-previewed on load — the big box stays empty
   until someone clicks a file. This keeps the page fast even
   when a department page has several of these carousels on it
   at once (Presentations, Documents, Spreadsheets, Training).
   ============================================================ */

interface ICarouselItem {
  name: string;
  isFolder: boolean;
  isPowerBiReport?: boolean; // NEW
  fileUrl?: string;
  folderServerRelativeUrl?: string;
  iconLabel: string;
  isPreviewableFile: boolean;
}

interface IBreadcrumbEntry {
  label: string;
  serverRelativeUrl: string;
}

interface IState {
  items: ICarouselItem[];
  breadcrumbs: IBreadcrumbEntry[];
  activeItem: ICarouselItem | undefined;
  isLoading: boolean;
  errorMessage: string;
  activeCategoryFilters: { [key in ItemCategory]: boolean };
  searchText: string;
  isSummaryPanelOpen: boolean; // NEW
}

type ItemCategory = 'ppt' | 'doc' | 'xls' | 'pdf' | 'bi' | 'other';

interface ICarouselItem {
  name: string;
  isFolder: boolean;
  isPowerBiReport?: boolean;
  fileUrl?: string;
  folderServerRelativeUrl?: string;
  iconLabel: string;
  isPreviewableFile: boolean;
  category: ItemCategory; // NEW - used for filtering
  uniqueId?: string; // NEW - stable file identity, used as a cache key for summaries
  etag?: string; // NEW - changes whenever the file's content changes
}

const PREVIEWABLE_EXTENSIONS = ['pptx', 'ppt', 'docx', 'doc', 'xlsx', 'xls', 'pdf'];

const FILE_TYPE_LABELS: { [extension: string]: string } = {
  pptx: 'PPT', ppt: 'PPT',
  docx: 'DOC', doc: 'DOC',
  xlsx: 'XLS', xls: 'XLS',
  pdf: 'PDF'
};

function getFileExtension(fileName: string): string {
  const parts = fileName.split('.');
  return parts[parts.length - 1].toLowerCase();
}

function getCategory(fileName: string, isPowerBiReport?: boolean): ItemCategory {
  if (isPowerBiReport) return 'bi';

  const ext = getFileExtension(fileName);
  if (ext === 'pptx' || ext === 'ppt') return 'ppt';
  if (ext === 'docx' || ext === 'doc') return 'doc';
  if (ext === 'xlsx' || ext === 'xls') return 'xls';
  if (ext === 'pdf') return 'pdf';
  return 'other';
}

function getFileTypeLabel(fileName: string): string {
  const ext = getFileExtension(fileName);
  return FILE_TYPE_LABELS[ext] || ext.toUpperCase().slice(0, 4) || 'FILE';
}

function canBePreviewedInline(fileName: string): boolean {
  return PREVIEWABLE_EXTENSIONS.indexOf(getFileExtension(fileName)) !== -1;
}

// Office files (docx/xlsx/pptx) need to go through SharePoint's WopiFrame
// viewer, since browsers can't render those formats natively - only
// Office's own rendering engine can. PDFs are different: browsers already
// know how to display PDFs on their own, so for those we just point the
// iframe straight at the file's own URL rather than routing through WOPI.
function buildPreviewUrl(item: ICarouselItem, siteAbsoluteUrl: string): string {
  if (item.isPowerBiReport && item.fileUrl) {
    return item.fileUrl; // Power BI's embed URL works directly in an iframe
  }

  const extension = getFileExtension(item.name);

  if (extension === 'pdf') {
    return item.fileUrl || '';
  }

  const encodedFileUrl = encodeURIComponent(item.fileUrl || '');
  return `${siteAbsoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodedFileUrl}&action=interactivepreview`;
}

// SharePoint's REST API wraps the folder path in single quotes inside the
// request URL. If the path itself contains a literal apostrophe, that would
// end the quoted string early and break the request. SharePoint's own
// convention is to escape a single apostrophe as TWO apostrophes ('').
function escapeApostrophesForODataQuery(path: string): string {
  if (typeof path !== 'string') {
    return '';
  }
  return path.replace(/'/g, "''");
}

// All inline styles live here, defined before the class that uses them.
const styles: { [key: string]: React.CSSProperties } = {
  container: {
    fontFamily: "'Segoe UI', sans-serif",
    maxWidth: '100%'
  },
  breadcrumbBar: {
    marginBottom: 10,
    fontSize: 13,
    color: '#605e5c'
  },
  breadcrumbLink: {
    color: '#0078d4',
    cursor: 'pointer'
  },
  breadcrumbCurrent: {
    color: '#323130',
    fontWeight: 600
  },
  breadcrumbSeparator: {
    color: '#a19f9d'
  },
  bigPreviewBox: {
    position: 'relative',
    width: '100%',
    height: 500,
    background: '#f3f2f1',
    borderRadius: 8,
    overflow: 'hidden',
    boxShadow: '0 2px 8px rgba(0,0,0,0.15)'
  },
  previewIframe: {
    width: '100%',
    height: '100%',
    border: 'none'
  },
  previewPlaceholder: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '100%',
    color: '#323130',
    padding: 20,
    textAlign: 'center'
  },
  previewUnsupportedLabel: {
    fontSize: 14,
    fontWeight: 600,
    color: '#0078d4',
    marginBottom: 10,
    letterSpacing: 1
  },
  previewUnsupportedHint: {
    fontSize: 13,
    color: '#0078d4',
    marginTop: 6
  },
  previewHint: {
    position: 'absolute',
    bottom: 10,
    right: 15,
    background: 'rgba(0,0,0,0.6)',
    color: 'white',
    padding: '6px 12px',
    borderRadius: 4,
    fontSize: 12,
    pointerEvents: 'none'
  },
  toolbarRow: {
    display: 'flex',
    gap: 10,
    marginTop: 14,
    alignItems: 'center'
  },
  searchInput: {
    flex: '1 1 auto',
    padding: '8px 10px',
    fontSize: 13,
    border: '1px solid #d1d1d1',
    borderRadius: 4
  },
  filterChipRow: {
    display: 'flex',
    gap: 8,
    marginTop: 10,
    flexWrap: 'wrap'
  },
  filterChip: {
    padding: '5px 12px',
    fontSize: 12,
    background: '#ffffff',
    color: '#605e5c',
    border: '1px solid #d1d1d1',
    borderRadius: 14,
    cursor: 'pointer'
  },
  filterChipActive: {
    padding: '5px 12px',
    fontSize: 12,
    background: '#0078d4',
    color: '#ffffff',
    border: '1px solid #0078d4',
    borderRadius: 14,
    cursor: 'pointer'
  },
  refreshButton: {
    padding: '8px 16px',
    fontSize: 13,
    background: '#ffffff',
    border: '1px solid #d1d1d1',
    borderRadius: 4,
    cursor: 'pointer'
  },
  summarizeButton: {
    padding: '8px 16px',
    fontSize: 13,
    background: '#0078d4',
    color: '#ffffff',
    border: '1px solid #0078d4',
    borderRadius: 4,
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: 6
  },
  summaryPanelLoading: {
    fontSize: 13,
    color: '#605e5c',
    padding: '14px 16px',
    marginTop: 12,
    marginBottom: 12,
    background: '#ffffff',
    border: '1px solid #d1d1d1',
    borderRadius: 8
  },
  carouselStrip: {
    display: 'flex',
    gap: 12,
    marginTop: 12,
    overflowX: 'auto',
    padding: '8px 4px'
  },
  thumbnailCard: {
    flex: '0 0 auto',
    width: 140,
    background: '#ffffff',
    borderRadius: 6,
    boxShadow: '0 1px 4px rgba(0,0,0,0.12)',
    cursor: 'pointer',
    textAlign: 'center',
    padding: '10px 8px'
  },
  thumbnailFileLabel: {
    fontSize: 14,
    fontWeight: 600,
    color: '#0078d4',
    marginBottom: 6,
    letterSpacing: 1
  },
  thumbnailFolderLabel: {
    fontSize: 14,
    fontWeight: 600,
    color: '#8a6d00',
    marginBottom: 6,
    letterSpacing: 1
  },
  thumbnailName: {
    fontSize: 12,
    color: '#323130',
    lineHeight: 1.3,
    wordBreak: 'break-word'
  },
  thumbnailPowerBiLabel: {
  fontSize: 14,
  fontWeight: 600,
  color: '#f2c811', // Power BI's brand yellow
  marginBottom: 6,
  letterSpacing: 1
  },
  statusMessage: {
    padding: 20,
    color: '#666'
  },
  errorMessage: {
    padding: 20,
    color: '#a80000'
  }
};

export default class DocumentPreviewCarousel extends React.Component<IDocumentPreviewCarouselProps, IState> {

  constructor(props: IDocumentPreviewCarouselProps) {
    super(props);
    this.state = {
      items: [],
      breadcrumbs: [],
      activeItem: undefined,
      isLoading: true,
      errorMessage: '',
      searchText: '',
      isSummaryPanelOpen: false,
      activeCategoryFilters: {
        ppt: props.defaultTypeFilters.showPpt,
        doc: props.defaultTypeFilters.showDoc,
        xls: props.defaultTypeFilters.showXls,
        pdf: props.defaultTypeFilters.showPdf,
        bi: props.defaultTypeFilters.showBi,
        other: props.defaultTypeFilters.showOther
      }
    };
  }

  public componentDidMount(): void {
    const startingFolder = this.props.folderServerRelativeUrl;
    if (startingFolder) {
      this.setState({
        breadcrumbs: [{ label: this.getLastFolderNameFromPath(startingFolder), serverRelativeUrl: startingFolder }]
      });
      this.loadFolder(startingFolder);
    } else {
      this.setState({ isLoading: false, errorMessage: 'No folder path configured yet. Edit this web part to set one.' });
    }
  }

  public componentDidUpdate(previousProps: IDocumentPreviewCarouselProps): void {
    if (previousProps.folderServerRelativeUrl !== this.props.folderServerRelativeUrl) {
      const newStartingFolder = this.props.folderServerRelativeUrl;
      this.setState({
        breadcrumbs: newStartingFolder
          ? [{ label: this.getLastFolderNameFromPath(newStartingFolder), serverRelativeUrl: newStartingFolder }]
          : [],
        searchText: ''
      });
      this.loadFolder(newStartingFolder);
    }
  }

  private getLastFolderNameFromPath(path: string): string {
    // Defensive check: if something unexpected got saved into the folder
    // path property (e.g. leftover bad data from an older version of this
    // web part), don't crash - just show something reasonable instead.
    if (typeof path !== 'string' || !path) {
      return 'Unknown folder';
    }
    const segments = path.split('/').filter((segment) => segment.length > 0);
    return segments[segments.length - 1] || path;
  }
  // Fetches ALL pages of results from a SharePoint REST API URL, not just
  // the first page. SharePoint typically caps a single response at around
  // 100 items and includes a "nextLink" URL to fetch the rest when there's
  // more. This keeps following that link until there's no more data.
  private fetchAllPages = (firstPageUrl: string): Promise<any[]> => {
    const { spHttpClient } = this.props;
    const allResults: any[] = [];

    const fetchPage = (url: string): Promise<any[]> => {
      return spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then((response: any) => {
          if (!response.ok) {
            const httpError: any = new Error('Request failed with status ' + response.status);
            httpError.statusCode = response.status;
            throw httpError;
          }
          return response.json();
        })
        .then((pageData: any) => {
          allResults.push(...(pageData.value || []));
          const nextLink = pageData['odata.nextLink'] || pageData['@odata.nextLink'];
          if (nextLink) {
            return fetchPage(nextLink);
          }
          return allResults;
        });
    };

    return fetchPage(firstPageUrl);
  }

  private loadFolder = (folderServerRelativeUrl: string): void => {
    const { siteAbsoluteUrl } = this.props;

    if (!folderServerRelativeUrl || typeof folderServerRelativeUrl !== 'string') {
      this.setState({ isLoading: false, errorMessage: 'No folder path configured yet. Edit this web part to set one.' });
      return;
    }

    this.setState({ isLoading: true, errorMessage: '', activeItem: undefined });

    const safePath = escapeApostrophesForODataQuery(folderServerRelativeUrl);
    // IMPORTANT: use encodeURI, not encodeURIComponent, here.
    // encodeURIComponent would also encode forward slashes (/ -> %2F),
    // which SharePoint's server rejects as invalid inside this part of
    // the URL, causing a 404 even though the folder genuinely exists.
    // encodeURI correctly encodes spaces but leaves / and & alone.
    const encodedPath = encodeURI(safePath);
    const origin = siteAbsoluteUrl.split('/').slice(0, 3).join('/');

    const filesRequestUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedPath}')/Files?$select=Name,ServerRelativeUrl,UniqueId,ETag&$top=500`;
    const foldersRequestUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedPath}')/Folders?$select=Name,ServerRelativeUrl&$top=500`;

    Promise.all([
      this.fetchAllPages(filesRequestUrl),
      this.fetchAllPages(foldersRequestUrl)
    ])
      .then(([allFiles, allFolders]: any[][]) => {

        const folderItems: ICarouselItem[] = allFolders
          .filter((folder: any) => folder.Name.indexOf('Forms') !== 0)
          .map((folder: any) => ({
            name: folder.Name,
            isFolder: true,
            folderServerRelativeUrl: folder.ServerRelativeUrl,
            iconLabel: 'FOLDER',
            isPreviewableFile: false,
            category: 'other' as ItemCategory // folders are always shown regardless of filter - see getFilteredItems
          }));
        const fileItems: ICarouselItem[] = allFiles.map((file: any) => ({
          name: file.Name,
          isFolder: false,
          fileUrl: origin + file.ServerRelativeUrl,
          iconLabel: getFileTypeLabel(file.Name),
          isPreviewableFile: canBePreviewedInline(file.Name),
          category: getCategory(file.Name),
          uniqueId: file.UniqueId,
          etag: file.ETag
        }));

        const powerBiItems: ICarouselItem[] = (this.props.powerBiReports || []).map((report) => ({
          name: report.name,
          isFolder: false,
          isPowerBiReport: true,
          fileUrl: report.embedUrl,
          iconLabel: 'BI',
          isPreviewableFile: true,
          category: getCategory(report.name, true)
        }));

        // Order: folders first (for navigation), then Power BI reports (pinned,
        // configured content), then regular files last.
        const allItems = folderItems.concat(powerBiItems).concat(fileItems);

        this.setState({
          items: allItems,
          activeItem: undefined, // don't auto-preview anything — wait for a click.
          isLoading: false
        });
      })
      .catch((error: any) => {
        if (error.statusCode === 403 || error.statusCode === 401) {
          this.setState({
            isLoading: false,
            errorMessage: "You don't have permission to view this folder. Contact your site owner if you think this is a mistake."
          });
        } else {
          this.setState({
            isLoading: false,
            errorMessage: 'Could not load this folder: ' + error.message
          });
        }
      });
  }

  private navigateIntoFolder = (folder: ICarouselItem): void => {
    if (!folder.folderServerRelativeUrl) return;

    const newBreadcrumb: IBreadcrumbEntry = {
      label: folder.name,
      serverRelativeUrl: folder.folderServerRelativeUrl
    };

    this.setState(
      (previousState) => ({
        breadcrumbs: previousState.breadcrumbs.concat([newBreadcrumb]),
        searchText: ''
      }),
      () => this.loadFolder(folder.folderServerRelativeUrl as string)
    );
  }

  private navigateToBreadcrumb = (clickedIndex: number): void => {
    this.setState(
      (previousState) => ({
        breadcrumbs: previousState.breadcrumbs.slice(0, clickedIndex + 1),
        searchText: ''
      }),
      () => {
        const targetFolder = this.state.breadcrumbs[this.state.breadcrumbs.length - 1];
        this.loadFolder(targetFolder.serverRelativeUrl);
      }
    );
  }

  private selectFile = (file: ICarouselItem): void => {
    this.setState({ activeItem: file });
  }

  private openActiveFileInNewTab = (): void => {
    if (this.state.activeItem && this.state.activeItem.fileUrl) {
      window.open(this.state.activeItem.fileUrl, '_blank');
    }
  }

  private toggleSummaryPanel = (): void => {
    this.setState((previousState) => ({ isSummaryPanelOpen: !previousState.isSummaryPanelOpen }));
  }

  private closeSummaryPanel = (): void => {
    this.setState({ isSummaryPanelOpen: false });
  }

  // Downloads the active file's raw bytes for parsing. Uses spHttpClient
  // (rather than a bare fetch()) so the request correctly carries the
  // SharePoint session context this web part already runs under - the same
  // client the rest of this component uses for REST calls.
  private fetchActiveFileBytes = (): Promise<ArrayBuffer> => {
    const { activeItem } = this.state;
    if (!activeItem || !activeItem.fileUrl) {
      return Promise.reject(new Error('No file selected.'));
    }
    return this.props.spHttpClient
      .get(activeItem.fileUrl, SPHttpClient.configurations.v1)
      .then((response: any) => {
        if (!response.ok) {
          throw new Error('Could not download this file (status ' + response.status + ').');
        }
        return response.arrayBuffer();
      });
  }

  // Builds what the SummaryPanel needs to know about the currently active
  // item, or undefined if nothing summarizable is selected right now - the
  // panel itself turns "undefined" into a clear fallback message rather
  // than crashing, per the agreed design (Power BI / no selection cases).
  private getSummaryTarget(): ISummaryTarget | undefined {
    const { activeItem } = this.state;
    if (!activeItem || activeItem.isFolder || activeItem.isPowerBiReport || !activeItem.fileUrl) {
      return undefined;
    }

    return {
      fileName: activeItem.name,
      // Falls back to the file URL if UniqueId/ETag are ever missing
      // (e.g. an older cached items list) rather than crashing - worst
      // case, the cache key is slightly less precise, not broken.
      fileId: activeItem.uniqueId || activeItem.fileUrl,
      fileVersion: activeItem.etag || 'unknown',
      fetchFileBytes: this.fetchActiveFileBytes
    };
  }

  private handleRefreshClick = (): void => {
    const currentFolder = this.state.breadcrumbs[this.state.breadcrumbs.length - 1];
    if (currentFolder) {
      this.loadFolder(currentFolder.serverRelativeUrl);
    }
  }

  private toggleCategoryFilter = (category: ItemCategory): void => {
    this.setState((previousState) => ({
    activeCategoryFilters: {
      ...previousState.activeCategoryFilters,
      [category]: !previousState.activeCategoryFilters[category]
    }
    }));
  }

  private setAllFiltersOn = (): void => {
    this.setState({
      activeCategoryFilters: { ppt: true, doc: true, xls: true, pdf: true, bi: true, other: true }
    });
  }

  private handleSearchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ searchText: event.target.value });
  }

  private getFilteredItems(): ICarouselItem[] {
    const { items, searchText, activeCategoryFilters } = this.state;

    const passesTypeFilter = (item: ICarouselItem): boolean => {
      if (item.isFolder) return true; // folders are never hidden by type filters - you need them to navigate
      return activeCategoryFilters[item.category];
    };

    const passesSearch = (item: ICarouselItem): boolean => {
      if (!searchText.trim()) return true;
      return item.name.toLowerCase().indexOf(searchText.toLowerCase()) !== -1;
    };

    return items.filter((item) => passesTypeFilter(item) && passesSearch(item));
  }

  private renderBigPreviewContents(): React.ReactNode {
    const { activeItem } = this.state;

    if (!activeItem) {
      return (
        <div style={styles.previewPlaceholder}>
          Select a file below to preview it here.
        </div>
      );
    }

    if (activeItem.isPreviewableFile && activeItem.fileUrl) {
      return (
        <iframe
          title="Document preview"
          style={styles.previewIframe}
          src={buildPreviewUrl(activeItem, this.props.siteAbsoluteUrl)}
        />
      );
    }

    return (
      <div style={styles.previewPlaceholder}>
        <div style={styles.previewUnsupportedLabel}>{activeItem.iconLabel}</div>
        <div>This file type can&apos;t be previewed here.</div>
        <div style={styles.previewUnsupportedHint}>Click below to open or download it instead.</div>
      </div>
    );
  }

  private renderBreadcrumbs(): React.ReactNode {
    const { breadcrumbs } = this.state;

    return (
      <div style={styles.breadcrumbBar}>
        {breadcrumbs.map((crumb, index) => {
          const isLastCrumb = index === breadcrumbs.length - 1;
          return (
            <React.Fragment key={index}>
              <span
                onClick={() => !isLastCrumb && this.navigateToBreadcrumb(index)}
                style={isLastCrumb ? styles.breadcrumbCurrent : styles.breadcrumbLink}
              >
                {crumb.label}
              </span>
              {!isLastCrumb && <span style={styles.breadcrumbSeparator}> / </span>}
            </React.Fragment>
          );
        })}
      </div>
    );
  }

  private renderToolbar(): React.ReactNode {
    return (
      <div style={styles.toolbarRow}>
        <input
          type="text"
          placeholder="Search this folder..."
          value={this.state.searchText}
          onChange={this.handleSearchChange}
          style={styles.searchInput}
        />
        <button onClick={this.toggleSummaryPanel} style={styles.summarizeButton}>
          Summarize AI
        </button>
        <button onClick={this.handleRefreshClick} style={styles.refreshButton}>
          Refresh
        </button>
      </div>
    );
  }

  private renderFilterChips(): React.ReactNode {
    const { activeCategoryFilters } = this.state;
    const allOn = Object.keys(activeCategoryFilters).every(
      (key) => activeCategoryFilters[key as ItemCategory]
    );

    const chipDefs: { category: ItemCategory; label: string }[] = [
      { category: 'ppt', label: 'PPT' },
      { category: 'doc', label: 'DOC' },
      { category: 'xls', label: 'XLS' },
      { category: 'pdf', label: 'PDF' },
      { category: 'bi', label: 'BI' },
      { category: 'other', label: 'Other' }
    ];

    return (
      <div style={styles.filterChipRow}>
        <button
          onClick={this.setAllFiltersOn}
          style={allOn ? styles.filterChipActive : styles.filterChip}
        >
          All
        </button>
        {chipDefs.map((chip) => (
          <button
            key={chip.category}
            onClick={() => this.toggleCategoryFilter(chip.category)}
            style={activeCategoryFilters[chip.category] ? styles.filterChipActive : styles.filterChip}
          >
            {chip.label}
          </button>
        ))}
      </div>
    );
  }

  private renderCarouselItem(item: ICarouselItem, index: number): React.ReactNode {
    const { activeItem } = this.state;
    const isSelected = !item.isFolder && activeItem !== undefined && item.fileUrl === activeItem.fileUrl;
  
    const labelStyle = item.isFolder
      ? styles.thumbnailFolderLabel
      : item.isPowerBiReport
        ? styles.thumbnailPowerBiLabel
        : styles.thumbnailFileLabel;
  
    return (
      <div
        key={index}
        onClick={() => item.isFolder ? this.navigateIntoFolder(item) : this.selectFile(item)}
        style={{
          ...styles.thumbnailCard,
          border: isSelected ? '2px solid #0078d4' : '2px solid transparent'
        }}
      >
        <div style={labelStyle}>
          {item.iconLabel}
        </div>
        <div style={styles.thumbnailName}>{item.name}</div>
      </div>
    );
  }

  public render(): React.ReactElement<IDocumentPreviewCarouselProps> {
    const { isLoading, errorMessage, activeItem } = this.state;
    const filteredItems = this.getFilteredItems();

    return (
      <div style={styles.container}>

        {this.renderBreadcrumbs()}

        <div
          onClick={this.openActiveFileInNewTab}
          style={{
            ...styles.bigPreviewBox,
            cursor: activeItem && activeItem.fileUrl ? 'pointer' : 'default'
          }}
        >
          {this.renderBigPreviewContents()}
          {activeItem && activeItem.fileUrl && (
            <div style={styles.previewHint}>
              {activeItem.isPreviewableFile ? 'Click to open full document' : 'Click to open or download'}
            </div>
          )}
        </div>

        {this.state.isSummaryPanelOpen && (
          <React.Suspense fallback={<div style={styles.summaryPanelLoading}>{'Loading summarizer\u2026'}</div>}>
            <SummaryPanel target={this.getSummaryTarget()} onClose={this.closeSummaryPanel} />
          </React.Suspense>
        )}

        {this.renderToolbar()}
        {this.renderFilterChips()}

        <div style={styles.carouselStrip}>
          {isLoading && <div style={styles.statusMessage}>Loading...</div>}
          {!isLoading && errorMessage && <div style={styles.errorMessage}>{errorMessage}</div>}
          {!isLoading && !errorMessage && filteredItems.length === 0 && (
            <div style={styles.statusMessage}>
              {this.state.searchText ? 'No items match your search.' : 'This folder is empty.'}
            </div>
          )}
          {!isLoading && filteredItems.map((item, index) => this.renderCarouselItem(item, index))}
        </div>

      </div>
    );
  }
}