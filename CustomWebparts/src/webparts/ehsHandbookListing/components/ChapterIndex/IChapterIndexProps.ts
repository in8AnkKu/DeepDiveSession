import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { HandbookComposite } from '../../../../BAL/HandbookComposite';

export interface IChapterIndexProps {
  handBookCompositeMetadata: HandbookComposite[];
  context: WebPartContext;
  selectedList: string;
  templateUrl: string;
  isContributor: boolean;
  isOwner: boolean;
  options: IDropdownOption[];
  includeChapterLinks: boolean;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
  sortOrder: string;
  sortType: string;
  additionalColumn: string[];
  selectetItem: string;
  sortByType: string;
  sortByField: string;
}