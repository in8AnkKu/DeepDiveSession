import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface ITrainingProps {
  context: WebPartContext;
  description: string;
  selectedList: string;
  displayMode: DisplayMode;
  configured: boolean;
  userAssessment: string;
  userTraining: string;
  moduleSubmittionMsg: string;
  moduleCompletionMsg: string;
  URLForYes: string;
  URLForNo: string;
}
