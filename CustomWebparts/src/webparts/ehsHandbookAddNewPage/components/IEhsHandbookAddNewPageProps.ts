import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface IEhsHandbookAddNewPageProps {
  pageType: string;
  context: WebPartContext;
  selectedList: string;
  templateUrl: string;
  configured: boolean;
  displayMode: DisplayMode;
  parentPageScope: string;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
  pageScopes: IDropdownOption[];
}
