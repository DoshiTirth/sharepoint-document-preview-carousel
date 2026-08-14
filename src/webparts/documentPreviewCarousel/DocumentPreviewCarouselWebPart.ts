import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import DocumentPreviewCarousel from './components/DocumentPreviewCarousel';
import { IDocumentPreviewCarouselProps, IPowerBiReportConfig } from './components/IDocumentPreviewCarouselProps';

export interface IDocumentPreviewCarouselWebPartProps {
  folderServerRelativeUrl: string;
  powerBiReportsRaw: string;
  showPptByDefault: boolean;
  showDocByDefault: boolean;
  showXlsByDefault: boolean;
  showPdfByDefault: boolean;
  showBiByDefault: boolean;
  showOtherByDefault: boolean;
}

export default class DocumentPreviewCarouselWebPart extends BaseClientSideWebPart<IDocumentPreviewCarouselWebPartProps> {

  // Turns the raw text someone typed into the property pane (one Power BI
  // report per line, formatted as "Name | Embed URL") into a clean list of
  // { name, embedUrl } objects that the carousel component can use directly.
  private parsePowerBiReports(raw: string): IPowerBiReportConfig[] {
    if (!raw) {
      return [];
    }

    return raw.split('\n')
      .map((line) => line.trim())
      .filter((line) => line.length > 0 && line.indexOf('|') !== -1)
      .map((line) => {
        const parts = line.split('|');
        return {
          name: parts[0].trim(),
          embedUrl: parts[1].trim()
        };
      });
  }

  public render(): void {
    const element: React.ReactElement<IDocumentPreviewCarouselProps> = React.createElement(
      DocumentPreviewCarousel,
      {
        folderServerRelativeUrl: this.properties.folderServerRelativeUrl,
        spHttpClient: this.context.spHttpClient,
        siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        powerBiReports: this.parsePowerBiReports(this.properties.powerBiReportsRaw),
        // !== false means these default to ON the very first time someone
        // adds this web part, before they've touched any settings at all
        // (an unset checkbox reads as "undefined", and undefined !== false
        // evaluates to true).
        defaultTypeFilters: {
          showPpt: this.properties.showPptByDefault !== false,
          showDoc: this.properties.showDocByDefault !== false,
          showXls: this.properties.showXlsByDefault !== false,
          showPdf: this.properties.showPdfByDefault !== false,
          showBi: this.properties.showBiByDefault !== false,
          showOther: this.properties.showOtherByDefault !== false
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Document Preview Carousel Settings' },
          groups: [
            {
              groupName: 'Folder & content',
              groupFields: [
                PropertyPaneTextField('folderServerRelativeUrl', {
                  label: 'Folder server-relative URL',
                  description: "e.g. /sites/YourSiteName/Shared Documents/Your Folder"
                }),
                PropertyPaneTextField('powerBiReportsRaw', {
                  label: 'Power BI reports (one per line: Name | Embed URL)',
                  multiline: true,
                  rows: 5,
                  description: 'Example: Sales Dashboard | https://app.powerbi.com/reportEmbed?...'
                })
              ]
            },
            {
              groupName: 'Default filters (viewers can still change these live)',
              groupFields: [
                PropertyPaneCheckbox('showPptByDefault', { text: 'Show PowerPoint files by default' }),
                PropertyPaneCheckbox('showDocByDefault', { text: 'Show Word documents by default' }),
                PropertyPaneCheckbox('showXlsByDefault', { text: 'Show Excel files by default' }),
                PropertyPaneCheckbox('showPdfByDefault', { text: 'Show PDF files by default' }),
                PropertyPaneCheckbox('showBiByDefault', { text: 'Show Power BI reports by default' }),
                PropertyPaneCheckbox('showOtherByDefault', { text: 'Show other file types by default' })
              ]
            }
          ]
        }
      ]
    };
  }
}