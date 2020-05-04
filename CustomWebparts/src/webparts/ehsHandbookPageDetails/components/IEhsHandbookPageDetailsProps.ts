import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IEhsHandbookPageDetailsProps {
  context: WebPartContext;
  layoutType: string;
  logoUrl: string;
  pageProperties: string[];
  configured: boolean;
  displayMode: DisplayMode;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
}