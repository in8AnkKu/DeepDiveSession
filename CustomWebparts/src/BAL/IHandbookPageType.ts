import { HandbookContentType } from './HandbookContentType';
import { HandbookColumn } from './HandbookColumn';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHandbookPageType {
  id: number;
  title: string;
  link: string;
  parentId: number;
  scope: string;
  bannerImageUrl: string;
  description: string;
  contentType?: HandbookContentType;
  getPageDetails?(selectedList: string, pageId: number, spContext: WebPartContext): Promise<IHandbookPageType>;
  getColumnsForContentType?(selectedList: string, pageId: number): HandbookColumn[];
}
