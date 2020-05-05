import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IAssessmentFinderProps {
  context: WebPartContext;
  description: string;
  selectedList: string;
  assessmentList:string;
  userAssessmentList: string;
}
