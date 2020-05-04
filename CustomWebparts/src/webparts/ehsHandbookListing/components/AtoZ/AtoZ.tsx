import * as React from 'react';
import EhsHandbookListingModuleScss from '../EhsHandbookListing.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { HandbookComposite } from '../../../../BAL/HandbookComposite';
import { HandbookLeaf } from '../../../../BAL/HandbookLeaf';

export interface IAtoZProps {
  sortAsc: boolean;
  context: WebPartContext;
  selectedList: string;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
  options: IDropdownOption[];
  handBookCompositeMetadata: HandbookComposite[];
  selectetItem: string;
  sortByField: string;
  additionalColumn: string[];
}

export interface IAtoZState {
  allItems: any[];
}

export class AtoZ extends React.Component<IAtoZProps, IAtoZState> {
  constructor(props: IAtoZProps) {
    super(props);

    this.state = {
      allItems: []
    };

    this.createElement = this.createElement.bind(this);
  }

  public componentDidUpdate(prevProps: IAtoZProps) {
    /* Render updated topics when the topicData property value is updated from EhsHandbookListing.tsx component*/
    if ((prevProps.handBookCompositeMetadata !== this.props.handBookCompositeMetadata) || (prevProps.sortAsc !== this.props.sortAsc) || (prevProps.additionalColumn !== this.props.additionalColumn) || (prevProps.sortByField !== this.props.sortByField)) {
      this.createElement();
    }
  }

  public componentDidMount() {
    this.createElement();
  }

  /**
   * Dynamically create element for rendering all the Topics in alphabetical order
   */
  private createElement() {
    try {
      let allTopics: HandbookLeaf[] = [];
      if (this.props.selectetItem === 'All') {
        this.props.handBookCompositeMetadata.map((currentSubject) => {
          if (currentSubject.childNodes !== undefined) {
            currentSubject.childNodes.map((currentChapter: HandbookComposite) => {
              if (currentChapter.childNodes !== undefined) {
                currentChapter.childNodes.map((currentTopic: HandbookLeaf) => {
                  allTopics.push(currentTopic);
                });
              }
            });
          }
        });
      } else {
        this.props.handBookCompositeMetadata.map((currentChapter) => {
          if (currentChapter.childNodes !== undefined) {
            currentChapter.childNodes.map((currentTopic: HandbookLeaf) => {
              allTopics.push(currentTopic);
            });
          }
        });
      }
      allTopics = this.props.sortAsc === true ? allTopics.sort((a, b) => a.title.localeCompare(b.title, 'en-u-kn-true')) : allTopics.sort((a, b) => b.title.localeCompare(a.title, 'en-u-kn-true'));
      let uniqueLetters = allTopics.map(topic => topic.title.charAt(0).toUpperCase()).filter((x, i, a) => a.indexOf(x) === i);

      let allLetters = [];
      let allItems = [];
      let cnt: number = 0;
      let additionalColumn: any = this.props.additionalColumn;
      let additionalColumnName: any = '';
      let sortOrder: any = this.props.sortByField === '' ? additionalColumn === 'none' ? 'title' : additionalColumn : this.props.sortByField;
      let months: string[] = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      uniqueLetters.map((currentUniqueLetter) => {
        if (currentUniqueLetter !== '') {
          cnt += 1;
          let currentLetterTopics = allTopics.filter(item => (item.title.charAt(0).toUpperCase() === currentUniqueLetter));
          let allTopicsArray = [];
          currentLetterTopics.sort(
            (item1, item2) => {
              let value1 = item1.contentType.columns[item1.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;
              let value2 = item2.contentType.columns[item2.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;
              return value1.localeCompare(value2, 'en-u-kn-true');
            })
            .map((currentTopic) => {
              let additionalField = additionalColumn === 'none' ? '' : currentTopic.contentType.columns.filter(column => column.internalName.toUpperCase() === additionalColumn.toUpperCase())[0].value;
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
                allTopicsArray.push(<div className={EhsHandbookListingModuleScss.topicLink}><a className={EhsHandbookListingModuleScss.pageLink} href={currentTopic.link}><span className={EhsHandbookListingModuleScss.additionalColumn}>{additionalField}</span>{currentTopic.title}</a> <Icon iconName='ProtectedDocument' /><br /></div>);
              } else {
                allTopicsArray.push(<div className={EhsHandbookListingModuleScss.topicLink}><a className={EhsHandbookListingModuleScss.pageLink} href={currentTopic.link}><span className={EhsHandbookListingModuleScss.additionalColumn}>{additionalField}</span>{currentTopic.title}</a></div>);
              }
            });

          allLetters.push(<div className={EhsHandbookListingModuleScss.indexcolumn}><div className={EhsHandbookListingModuleScss.chapterHeading}>{currentUniqueLetter}</div>{allTopicsArray}</div>);
          if (cnt % 3 === 0) {
            allItems.push(<div className={EhsHandbookListingModuleScss.indexRow}>{allLetters[cnt - 3]}{allLetters[cnt - 2]}{allLetters[cnt - 1]}</div>);
          }
        }
      });

      if (cnt % 3 === 1) {
        allItems.push(<div className={EhsHandbookListingModuleScss.indexRow}>{allLetters[cnt - 1]}</div>);
      } else if (cnt % 3 === 2) {
        allItems.push(<div className={EhsHandbookListingModuleScss.indexRow}>{allLetters[cnt - 2]}{allLetters[cnt - 1]}</div>);
      }
      this.setState({ allItems });
    } catch (error) {
      console.log({ errorMessage: error.message, errorMethod: 'AtoZ.createElement' });
    }
  }

  public render(): React.ReactElement<{}> {
    return (
      <div>
        {this.state.allItems}
      </div>
    );
  }
}
