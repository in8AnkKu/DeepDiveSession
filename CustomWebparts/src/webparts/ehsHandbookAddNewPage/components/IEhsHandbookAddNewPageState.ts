import { IDropdownOption } from 'office-ui-fabric-react';

export interface IEhsHandbookAddNewPageState {
  showPanel: boolean;
  pageName: string;
  topicPageLayout: string;
  pageNameErrorMessage: string;
  pageTemplateUrl: string;
  subjectDescription: string;
  subjectImage: string;
  subjectImageErrorText: string;
  subjectBannerImage: string;
  subjectBannerImageErrorText: string;
  loading: boolean;
  scope: IDropdownOption;
}