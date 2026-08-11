export interface IPowerBiReportConfig {
  name: string;
  embedUrl: string;
}

// Which file types should be visible by default when the carousel first
// loads. Viewers can still change this live using the filter chips in the
// carousel itself - this just controls the STARTING point.
export interface IDefaultTypeFilters {
  showPpt: boolean;
  showDoc: boolean;
  showXls: boolean;
  showPdf: boolean;
  showBi: boolean;
  showOther: boolean; // anything that isn't one of the above (zip, image, etc.)
}

export interface IDocumentPreviewCarouselProps {
  folderServerRelativeUrl: string;
  spHttpClient: any;
  siteAbsoluteUrl: string;
  powerBiReports: IPowerBiReportConfig[];
  defaultTypeFilters: IDefaultTypeFilters;
}