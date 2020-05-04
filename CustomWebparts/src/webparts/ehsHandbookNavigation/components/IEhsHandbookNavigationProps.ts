import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IEhsHandbookNavigationProps {
  context: WebPartContext;
  selectedList: string;
  configured: boolean;
  displayMode: DisplayMode;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
}
