import { IDropdownOption} from 'office-ui-fabric-react/lib/Dropdown';

export interface IChapterIndexState {
  showPanel: boolean;
  pageTemplateUrl: string;
  pageName: string;
  topicPageLayout: string;
  pageNameErrorMessage: string;
  topicLoading: boolean;
  isTopicError: boolean;
  topicLayoutError: string;
  subjectDescription: string;
  subjectImage: string;
  subjectBannerImage: string;
  loading: boolean;
  allItems: any[];
  scope: IDropdownOption;
}