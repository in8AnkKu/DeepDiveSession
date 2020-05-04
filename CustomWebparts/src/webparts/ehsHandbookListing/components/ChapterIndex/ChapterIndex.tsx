import * as React from 'react';
import EhsHandbookListingModuleScss from '../EhsHandbookListing.module.scss';
import {
  Dropdown,
  IDropdownOption,
  Icon,
  TextField,
  DefaultButton,
  PrimaryButton,
  Panel, Spinner
} from 'office-ui-fabric-react';
import { IChapterIndexProps } from './IChapterIndexProps';
import { IChapterIndexState } from './IChapterIndexState';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { v4 } from 'uuid';
import { unstable_renderSubtreeIntoContainer } from 'react-dom';
import { HandbookComposite } from '../../../../BAL/HandbookComposite';

export class ChapterIndex extends React.Component<IChapterIndexProps, IChapterIndexState> {
  /* Id of the Parent Chapter page to be used in adding a new topic */
  private parentPageId: number;
  private parentPageScope: string;
  private static defaultIcon: string;
  private static topic: any[];
  private static topicPageLayoutValue: string;
  private allOptions: IDropdownOption[];

  constructor(props: IChapterIndexProps) {
    super(props);
    this.getTopicTemplate();

    this.state = {
      showPanel: false,
      pageTemplateUrl: '',
      pageName: '',
      topicPageLayout: '',
      pageNameErrorMessage: '',
      isTopicError: false,
      subjectDescription: '',
      subjectImage: '',
      subjectBannerImage: '',
      loading: false,
      allItems: [],
      topicLayoutError: '',
      topicLoading: true,
      scope: undefined
    };
    this.createElement = this.createElement.bind(this);
    this.addNewPage = this.addNewPage.bind(this);
    this.validateFields = this.validateFields.bind(this);
    this.addTopicOnClick = this.addTopicOnClick.bind(this);
    this.onChange = this.onChange.bind(this);
  }

  public componentDidUpdate(prevProps: IChapterIndexProps) {

    /* Render updated topics when the topicData property value is updated from EhsHandbookListing.tsx component*/
    if ((prevProps.handBookCompositeMetadata !== this.props.handBookCompositeMetadata) || (prevProps.includeChapterLinks !== this.props.includeChapterLinks) || (prevProps.sortOrder !== this.props.sortOrder) || (prevProps.sortType !== this.props.sortType) || (prevProps.additionalColumn !== this.props.additionalColumn) || (prevProps.sortByType !== this.props.sortByType) || (prevProps.sortByField !== this.props.sortByField)) {
      this.createElement();
    }
    if (prevProps.options !== this.props.options) {
      this.allOptions = this.props.options;
    }
  }

  public componentDidMount() {

    this.allOptions = this.props.options;
    this.createElement();
  }

  /**
* This property sets the value of topicPageLayout to the selected topic template ID
*/
  private onChange(ev: React.FormEvent<HTMLInputElement>, option: any) {

    ChapterIndex.topicPageLayoutValue = option.key;
    this.setState({ topicPageLayout: ChapterIndex.topicPageLayoutValue });

  }
  /**
   * Gets all the topic templates from the site
   */
  private async getTopicTemplate() {
    let topicTemplateOptions = [];
    const uuid = require('uuid/v4');

    try {
      let handbookComposite = new HandbookComposite(this.props.context);
      let pageData = await handbookComposite.getPageData(this.props.context.pageContext.list.id.toString());
      console.log(pageData);
      if (pageData.length === 0) {
        this.setState({ isTopicError: true });
        this.setState({ topicLoading: false });
      } else {
        pageData.forEach(topicPage => {
          ChapterIndex.defaultIcon = this.props.context.pageContext.web.absoluteUrl + '/SiteAssets/placeholder-image.png?' + uuid();
          topicTemplateOptions.push({ key: topicPage.ID, text: topicPage.Title, imageSrc: topicPage.TopicIcon == null ? ChapterIndex.defaultIcon : topicPage.TopicIcon.Url, selectedImageSrc: topicPage.SelectedTopicIcon == null ? ChapterIndex.defaultIcon : topicPage.SelectedTopicIcon.Url });
        });
        ChapterIndex.topic = topicTemplateOptions;
        this.setState({ topicPageLayout: ChapterIndex.topic[0].key });
        this.setState({ topicLoading: false });

      }
    } catch (error) {
      console.log({ errorMessage: error.message, errorMethod: 'HandbookAddNewPage.getTopicTemplate' });
    }

  }

  private addTopicOnClick(currentChapterId: number, currentChapterScope: string) {
    this.allOptions = this.props.options;
    this.parentPageId = currentChapterId;
    this.parentPageScope = currentChapterScope;
    this.setState({ showPanel: true, scope: this.props.options.filter(pageScope => pageScope.key === this.parentPageScope)[0] });
    if (this.allOptions.length > 0) {
      switch (this.parentPageScope) {
        case this.allOptions[0] && this.allOptions[0].key: //External
          this.allOptions[0].disabled = false;
          this.allOptions[1].disabled = false;
          break;
        case this.allOptions[1] && this.allOptions[1].key: //Internal
          this.allOptions[0].disabled = true;
          this.allOptions[1].disabled = false;
          break;
        case this.allOptions[2] && this.allOptions[2].key:
          this.allOptions[0].disabled = true;
          this.allOptions[1].disabled = true;
          break;
        default:
          break;
      }
    }
  }

  /**
   * Dynamically create element for rendering all the Chapters and Topics
   */
  private createElement() {
    try {
      let subjectAllChapters = [];
      let sortOrder: any = this.props.sortByField === '' ? this.props.sortOrder : this.props.sortByField;
      let sortType: any = this.props.sortByType === '' ? this.props.sortType : this.props.sortByType;
      let additionalColumn: any = this.props.additionalColumn;
      let additionalColumnName: any = '';
      let months: string[] = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      if (this.props.selectetItem === 'All') {
        let sortedPageMetadata: any = this.sortData(sortType, sortOrder, this.props.handBookCompositeMetadata);
        sortedPageMetadata.filter((sub) => {
          return (sub.childNodes.length !== 0);
        }).map((currentSubject) => {
          this.getSubject(currentSubject, sortType, sortOrder, additionalColumn, additionalColumnName, months, subjectAllChapters);
        });
      } else {
        this.getSubject(this.props.handBookCompositeMetadata[0], sortType, sortOrder, additionalColumn, additionalColumnName, months, subjectAllChapters);
      }
      this.setState({ allItems: subjectAllChapters });
    } catch (error) {
      console.log({ errorMessage: error.message, errorMethod: 'ChapterIndex.createElement' });
    }
  }

  private getSubject(currentSubject: HandbookComposite, sortType: any, sortOrder: any, additionalColumn: any, additionalColumnName: any, months: string[], subjectAllChapters: any[]) {
    let subjectChapters = [];
    let columnCount = 0;
    if (this.props.selectetItem === 'All') {
      if (currentSubject.childNodes !== undefined) {
        let sortedCurrentSubject: any = this.sortData(sortType, sortOrder, currentSubject.childNodes);
        sortedCurrentSubject.map((currentChapter: HandbookComposite) => {
          columnCount = this.getTopic(columnCount, currentChapter, sortType, sortOrder, additionalColumn, additionalColumnName, months, subjectChapters, subjectAllChapters);
        });
      }
    } else {
      columnCount = this.getTopic(columnCount, currentSubject, sortType, sortOrder, additionalColumn, additionalColumnName, months, subjectChapters, subjectAllChapters);
    }
    if (columnCount % 3 === 1) {
      subjectAllChapters.push(<div className={EhsHandbookListingModuleScss.indexRow}>{subjectChapters[columnCount - 1]}</div>);
    } else if (columnCount % 3 === 2) {
      subjectAllChapters.push(<div className={EhsHandbookListingModuleScss.indexRow}>{subjectChapters[columnCount - 2]}{subjectChapters[columnCount - 1]}</div>);
    }
  }

  private getTopic(columnCount: number, currentChapter: HandbookComposite, sortType: any, sortOrder: any, additionalColumn: any, additionalColumnName: any, months: string[], subjectChapters: any[], subjectAllChapters: any[]) {
    columnCount += 1;
    let chapterTitle = currentChapter.title;
    let chapterLink = currentChapter.link;
    let chapterTopics = [];
    if (currentChapter.childNodes !== undefined) {
      let sortedCurrentChapter: any = this.sortData(sortType, sortOrder, currentChapter.childNodes);
      sortedCurrentChapter.map((currentTopic) => {
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
          chapterTopics.push(<div className={EhsHandbookListingModuleScss.topicLink}><a className={EhsHandbookListingModuleScss.pageLink} href={currentTopic.link}><span className={EhsHandbookListingModuleScss.additionalColumn}>{additionalField}</span>{currentTopic.title}</a> <Icon iconName='ProtectedDocument' /><br /></div>);
        } else {
          chapterTopics.push(<div className={EhsHandbookListingModuleScss.topicLink}><a className={EhsHandbookListingModuleScss.pageLink} href={currentTopic.link}><span className={EhsHandbookListingModuleScss.additionalColumn}>{additionalField}</span>{currentTopic.title}</a></div>);
        }
      });
    }
    if (this.props.isOwner === true) {
      chapterTopics.push(<div className={EhsHandbookListingModuleScss.topicLink}><DefaultButton text='Add Topic' onClick={() => this.addTopicOnClick(currentChapter.id, currentChapter.scope)} iconProps={{ iconName: 'CalculatorAddition', styles: { root: EhsHandbookListingModuleScss.calculatorAdditionIcon } }} styles={{ root: { padding: '10px', marginTop: '10px' } }} /></div>);
    } else if (this.props.isContributor === true) {
      chapterTopics.push(<div className={EhsHandbookListingModuleScss.topicLink}><DefaultButton text='Add Topic' onClick={() => this.addTopicOnClick(currentChapter.id, currentChapter.scope)} iconProps={{ iconName: 'CalculatorAddition', styles: { root: EhsHandbookListingModuleScss.calculatorAdditionIcon } }} styles={{ root: { padding: '10px', marginTop: '10px' } }} /></div>);
    }
    if (this.props.includeChapterLinks === true) {
      subjectChapters.push(<div className={EhsHandbookListingModuleScss.indexcolumn}><div className={EhsHandbookListingModuleScss.chapterHeading}><a className={EhsHandbookListingModuleScss.pageLink} href={chapterLink}>{chapterTitle}</a></div>{chapterTopics}</div>);
    } else {
      subjectChapters.push(<div className={EhsHandbookListingModuleScss.indexcolumn}><div className={EhsHandbookListingModuleScss.chapterHeading}>{chapterTitle}</div>{chapterTopics}</div>);
    }
    if (columnCount % 3 === 0) {
      subjectAllChapters.push(<div className={EhsHandbookListingModuleScss.indexRow}>{subjectChapters[columnCount - 3]}{subjectChapters[columnCount - 2]}{subjectChapters[columnCount - 1]}</div>);
    }
    return columnCount;
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

  /**
   * Validation check for adding a new topic
   */
  private validateFields() {
    if (this.state.pageName === '') {
      this.setState({ pageNameErrorMessage: 'Please enter the Topic Name' });
    } else {
      this.addNewPage();
    }
  }

  /**
   * Update the Spinner visibility and Add the new Topic page
   */
  private async addNewPage() {
    let compositeContext = new HandbookComposite(this.props.context);
    this.setState({ loading: true });
    compositeContext.createNewPage(this.state.pageName, this.state.topicPageLayout, this.parentPageId, this.props, this.state, this.state.scope.key, this.props.selectedList);
  }

  /**
   * Footer content for Add a Topic Panel
   */
  private onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this.validateFields} style={{ marginRight: '8px' }}>
          Save
        </PrimaryButton>
        <DefaultButton onClick={() => this.setState({ showPanel: false, pageName: '', subjectDescription: '', subjectImage: '', subjectBannerImage: '', scope: undefined })}>Cancel</DefaultButton>
      </div>
    );
  }

  public render(): React.ReactElement<{}> {
    return (
      <div>
        {this.state.allItems}
        < Panel headerText='Add Topic'
          isOpen={this.state.showPanel}
          onDismiss={() => this.setState({ showPanel: false, pageName: '', subjectDescription: '', subjectImage: '', subjectBannerImage: '', scope: undefined })}
          onRenderFooterContent={this.onRenderFooterContent}
        >
          <Dropdown label='Scope'
            defaultSelectedKey={this.state.scope ? this.state.scope.key : undefined}
            options={this.allOptions}
            onChanged={(val) => { this.setState({ scope: val }); }}
          />
          <TextField required={true} errorMessage={this.state.pageNameErrorMessage} label='Topic Name' value={this.state.pageName} onChanged={(value) => { this.setState({ pageName: value }); }} />
          {
            this.state.isTopicError === false &&
            <div>
              <ChoiceGroup
                label={'Topic Page Layout'}
                defaultSelectedKey={this.state.topicPageLayout}
                options={ChapterIndex.topic}
                onChange={this.onChange}
                required={true}
              />
              <br />
              {this.state.topicLoading === true &&
                <Spinner label='Getting All Topic Templates' ariaLive='assertive' />
              }
            </div>
          }
          <br />
          {
            this.state.loading === true &&
            <Spinner label='Creating New Topic' ariaLive='assertive' />
          }
        </Panel >
      </div >
    );
  }
}
