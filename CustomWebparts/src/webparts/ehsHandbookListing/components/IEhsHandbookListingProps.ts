import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IEhsHandbookListingProps {
  webPartView: string;
  selectedSibling: string;
  selectedChapter: string;
  context: WebPartContext;
  selectedList: string;
  includeChapterLinks: boolean;
  topicTemplateUrl: string;
  isContributor: boolean;
  isOwner: boolean;
  configured: boolean;
  displayMode: DisplayMode;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
  filters: string[];
  allTermSetFields: any[];
  sortByColumnName: string;
  sortByAscOrDesc: string;
  additionalField: string[];
  showChildOrGrandChild: string;
  defaultShowChildCount: number;
}
