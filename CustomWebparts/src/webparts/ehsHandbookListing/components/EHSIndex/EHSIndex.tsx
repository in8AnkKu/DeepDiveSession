import * as React from 'react';
import EhsHandbookListingModuleScss from '../EhsHandbookListing.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, IDropdownOption } from 'office-ui-fabric-react';
import { HandbookComposite } from '../../../../BAL/HandbookComposite';

export interface IEHSIndexProps {
  context: WebPartContext;
  selectedList: string;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
  options: IDropdownOption[];
  includeChapterLinks: boolean;
  sortOrder: string;
  sortType: string;
  additionalColumn: string[];
  handBookCompositeMetadata: HandbookComposite[];
  sortByType: string;
  sortByField: string;
}

export interface IEHSIndexState {
  allItems: any[];
}

export class EHSIndex extends React.Component<IEHSIndexProps, IEHSIndexState> {
  constructor(props: IEHSIndexProps) {
    super(props);

    this.createElement = this.createElement.bind(this);

    this.state = {
      allItems: []
    };
  }

  public componentDidMount() {
    this.createElement();
  }

  public componentDidUpdate(prevProps: IEHSIndexProps) {
    /* Render updated topics when the pageMetadata property value is updated from EhsHandbookListing.tsx component*/
    if ((prevProps.handBookCompositeMetadata !== this.props.handBookCompositeMetadata) || (prevProps.includeChapterLinks !== this.props.includeChapterLinks) || (prevProps.sortOrder !== this.props.sortOrder) || (prevProps.sortType !== this.props.sortType) || (prevProps.additionalColumn !== this.props.additionalColumn) || (prevProps.sortByType !== this.props.sortByType) || (prevProps.sortByField !== this.props.sortByField)) {
      this.createElement();
    }
  }

  /**
   * Dynamically create element for rendering all the Subjects, Chapters and Topics
   */
  private createElement() {
    try {
      let allSubjectsTopics = [];
      let sortOrder: any = this.props.sortByField === '' ? this.props.sortOrder : this.props.sortByField;
      let sortType: any = this.props.sortByType === '' ? this.props.sortType : this.props.sortByType;
      let additionalColumn: any = this.props.additionalColumn;
      let additionalColumnName: any = '';
      let months: string[] = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      let sortedPageMetadata: any = this.sortData(sortType, sortOrder, this.props.handBookCompositeMetadata);
      sortedPageMetadata.filter((sub) => { return 'childNodes' in sub; }).filter((sub) => { return (sub.childNodes.length !== 0); }).map((currentSubject) => {
        let subjectTitle = currentSubject.title;
        let subjectChapters = [];
        let subjectAllChapters = [];
        let columnCount = 0;
        if (currentSubject.childNodes !== undefined) {
          let sortedCurrentSubject: any = this.sortData(sortType, sortOrder, currentSubject.childNodes);
          sortedCurrentSubject.filter((chapter: HandbookComposite) => { if (chapter.childNodes !== undefined) { return (chapter.childNodes.length !== 0); } }).map((currentChapter: HandbookComposite) => {
            columnCount += 1;
            let chapterTitle = currentChapter.title;
            let chapterLink = currentChapter.link;
            let chapterTopics = [];
            if (currentChapter.childNodes !== undefined) {
              let sortedCurrentChapter: any = this.sortData(sortType, sortOrder, currentChapter.childNodes);
              sortedCurrentChapter.map((currentTopic) => {
                let additionalField = additionalColumn; // === 'none' ? '' : currentTopic.contentType.columns.filter(column => column.internalName.toUpperCase() === additionalColumn.toUpperCase())[0].value;
                if (additionalColumn === 'created' || additionalColumn === 'modified') {
                  additionalField = new Date(additionalField);
                  let dateMonth = additionalField.getMonth();
                  let dateDay = additionalField.getDate();
                  let dateYear = additionalField.getFullYear();
                  additionalField = months[dateMonth] + ' ' + dateDay + ', ' + dateYear;
                  additionalColumnName = additionalColumn === 'created' ? 'Created' : 'Modified';
                  additionalColumnName = additionalColumnName + ': ';
                }
                if (additionalColumn === 'fileLeafRef') {
                  additionalColumnName = 'Name: ';
                  additionalField = additionalField === null ? additionalField : additionalField.substr(0, additionalField.lastIndexOf('.'));
                }
                additionalField = additionalColumnName + additionalField;
                if ((this.props.options[1] && (currentTopic.scope === this.props.options[1].key)) || (this.props.options[2] && (currentTopic.scope === this.props.options[2].key))) {
                  chapterTopics.push(<div className={EhsHandbookListingModuleScss.topicLink}><a className={EhsHandbookListingModuleScss.pageLink} href={currentTopic.link}><span className={EhsHandbookListingModuleScss.additionalColumn}>{additionalField}</span>{currentTopic.title}</a> <Icon iconName='ProtectedDocument' /><br /></div>);
                } else {
                  chapterTopics.push(<div className={EhsHandbookListingModuleScss.topicLink}><a className={EhsHandbookListingModuleScss.pageLink} href={currentTopic.link}><span className={EhsHandbookListingModuleScss.additionalColumn}>{additionalField}</span>{currentTopic.title}</a></div>);
                }
              });
            }
            if (chapterTopics.length !== 0) {
              if (this.props.includeChapterLinks === true) {
                subjectChapters.push(<div className={EhsHandbookListingModuleScss.indexcolumn}><div className={EhsHandbookListingModuleScss.chapterHeading}><a className={EhsHandbookListingModuleScss.pageLink} href={chapterLink}>{chapterTitle}</a></div>{chapterTopics}</div>);
              } else {
                subjectChapters.push(<div className={EhsHandbookListingModuleScss.indexcolumn}><div className={EhsHandbookListingModuleScss.chapterHeading}>{chapterTitle}</div>{chapterTopics}</div>);
              }
              if (columnCount % 3 === 0) {
                subjectAllChapters.push(<div className={EhsHandbookListingModuleScss.indexRow}>{subjectChapters[columnCount - 3]}{subjectChapters[columnCount - 2]}{subjectChapters[columnCount - 1]}</div>);
              }
            }
          });
        }
        if (columnCount % 3 === 1) {
          subjectAllChapters.push(<div className={EhsHandbookListingModuleScss.indexRow}>{subjectChapters[columnCount - 1]}</div>);
        } else if (columnCount % 3 === 2) {
          subjectAllChapters.push(<div className={EhsHandbookListingModuleScss.indexRow}>{subjectChapters[columnCount - 2]}{subjectChapters[columnCount - 1]}</div>);
        }
        if (subjectChapters.length !== 0) {
          allSubjectsTopics.push(<div><div className={EhsHandbookListingModuleScss.subHeader}>{subjectTitle}<div className={EhsHandbookListingModuleScss.subjectLine}></div></div>{subjectAllChapters}<div className={EhsHandbookListingModuleScss.subjectDivider}></div></div>);
        }
      });
      this.setState({ allItems: allSubjectsTopics });
    } catch (error) {
      console.log({ errorMessage: error.message, errorMethod: 'EHSIndex.createElement' });
    }
  }

  private sortData(sortType: any, sortOrder: any, handbookItems: any): any {
    return sortType === 'Asc' ?
      handbookItems.sort((item1, item2) => {
        let value1 = item1.contentType.columns[item1.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;
        let value2 = item2.contentType.columns[item2.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;

        return value1.localeCompare(value2, 'en-u-kn-true');
      })
      : handbookItems.sort((item1, item2) => {
        let value1 = item1.contentType.columns[item1.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;
        let value2 = item2.contentType.columns[item2.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;
        return value2.localeCompare(value1, 'en-u-kn-true');
      });
  }

  public render(): React.ReactElement<{}> {
    return (
      <div>
        {this.state.allItems}
      </div >
    );
  }
}