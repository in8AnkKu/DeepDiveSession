import * as React from 'react';
import EhsHandbookListingModuleScss from './EhsHandbookListing.module.scss';
import { IEhsHandbookListingProps } from './IEhsHandbookListingProps';
import { IEhsHandbookListingState } from './IEhsHandbookListingState';
import {
  Panel,
  Pivot, PivotItem,
  PrimaryButton, DefaultButton,
  Icon,
  IContextualMenuItem, ContextualMenuItemType,
  IDropdownOption
} from 'office-ui-fabric-react';
import { TaxonomyPicker, IPickerTerms, Placeholder } from '@pnp/spfx-controls-react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { Subjects } from './Subjects/Subjects';
import { EHSIndex } from './EHSIndex/EHSIndex';
import { ChapterIndex } from './ChapterIndex/ChapterIndex';
import { AtoZ } from './AtoZ/AtoZ';
import { HandbookComposite } from '../../../BAL/HandbookComposite';
import { HandbookColumn } from '../../../BAL/HandbookColumn';
import { HandbookLeaf } from '../../../BAL/HandbookLeaf';
import { HandbookContentType } from '../../../BAL/HandbookContentType';
import * as strings from 'EhsHandbookListingWebPartStrings';
import { IHandbookContentType } from '../../../BAL/IHandbookContentType';

export default class EhsHandbookListing extends React.Component<IEhsHandbookListingProps, IEhsHandbookListingState> {
  /* Topics data stored in global variable to be used in filtering*/
  private handBookCompositeMetadata: HandbookComposite[] = [];
  /* Initial filter state to be used in filtering*/
  private initialFilterState: { name: string, internalName: string, termsetId: string, id: string, allowMultipleValues: boolean, value: IPickerTerms }[] = [];
  private currentPageID: number = this.props.context.pageContext.listItem.id;
  private parentID: number = null;
  private contentTypeLevel: number = null;
  private contentTypeDetail: IHandbookContentType = {
    name: '',
    contentTypeID: '',
    parentContentTypeID: '',
    parentContentTypeLength: 0,
    contentTypeLevel: 0,
    description: ''
  };

  constructor(props: IEhsHandbookListingProps) {
    super(props);

    this.handleLinkClick = this.handleLinkClick.bind(this);
    this.applyFilter = this.applyFilter.bind(this);
    this.onConfigure = this.onConfigure.bind(this);
    this.handleClick = this.handleClick.bind(this);

    this.state = {
      showPanel: false,
      hideDialog: true,
      pageName: '',
      selectedKey: this.props.webPartView === strings.siblingListingOptionKey ? strings.siblingListingOptionKey : 'ChapterIndex',
      sortText: 'A to Z',
      sortIcon: 'SortUp',
      sortAsc: true,
      filterDisplay: ((this.props.webPartView === strings.siblingListingOptionKey) || (this.props.filters.length === 0)) ? 'none' : 'block',
      handBookCompositeMetadata: [],
      siblingMetadata: [],
      statusText: '',
      filterState: [],
      options: [],
      sortByField: this.props.sortByColumnName,
      sortByType: this.props.sortByAscOrDesc,
      pageContentType: ''
    };

    if ((this.props.filters.length !== 0) && (this.props.allTermSetFields.length !== 0)) {
      this.props.filters.map((filter) => {
        let termSetFieldId = filter.substr(filter.length - 7) === 'TermSet' ? filter.substr(0, filter.length - 7) : filter;
        let termSetFieldValue = filter.substr(filter.length - 7) === 'TermSet' ? 'TermSetId' : 'AnchorId';
        let field = this.props.allTermSetFields.filter(termSetField => termSetField[termSetFieldValue] === termSetFieldId)[0];
        if (field !== undefined) {
          this.state.filterState.push({ name: field.Title, internalName: field.InternalName, termsetId: field.TermSetId, id: field[termSetFieldValue], allowMultipleValues: field.AllowMultipleValues, value: [] });
          this.initialFilterState.push({ name: field.Title, internalName: field.InternalName, termsetId: field.TermSetId, id: field[termSetFieldValue], allowMultipleValues: field.AllowMultipleValues, value: [] });
        }
      });
    }
  }

  public async componentDidMount() {
    /* Get all the Handbook Siblings and Topics data if Sibling view is selected in the web part
    else get all Chapters and Topics for the selected sibling if Chapter view is selected*/
    let compositeNode = new HandbookComposite(this.props.context);
    let handbookContentType = new HandbookContentType(this.props.context);
    if ((this.props.webPartView === strings.siblingListingOptionKey) && (this.props.selectedList !== undefined)) {
      let contentTypes = await handbookContentType.populateContentTypeLevel(this.props.selectedList);
      let pageDetail = await compositeNode.getPageDetailsById(escape(this.props.selectedList), this.currentPageID);
      contentTypes.filter((contentType) => {
        if (contentType.contentTypeID === pageDetail.ContentTypeId) {
          this.contentTypeDetail.contentTypeID = contentType.contentTypeID;
          this.contentTypeDetail.contentTypeLevel = contentType.contentTypeLevel;
          this.contentTypeDetail.name = contentType.name;
        }
      });
      if (this.contentTypeDetail.contentTypeID === '') {
        this.contentTypeDetail.contentTypeID = contentTypes[0].contentTypeID;
        this.contentTypeDetail.contentTypeLevel = contentTypes[0].contentTypeLevel;
        this.contentTypeDetail.name = contentTypes[0].name;
      }
      this.parentID = pageDetail.ParentId;
      if (pageDetail.ParentId === null && this.contentTypeDetail.contentTypeID.indexOf(pageDetail.ContentTypeId) === -1) {
        this.parentID = null;
      }
      this.contentTypeLevel = this.contentTypeDetail.contentTypeLevel + 1;

      let topicData: any = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), this.parentID, null, this.props.context);
      this.handBookCompositeMetadata = topicData;
      if (topicData.length === 0) {
        this.setState({ handBookCompositeMetadata: topicData, statusText: 'No Chapters or Topics have been published' });
      } else {
        if (this.parentID !== null && topicData.length > 0) {
          this.setState({ handBookCompositeMetadata: topicData, siblingMetadata: topicData[0].childNodes || [], pageContentType: this.contentTypeDetail.name });
        } else {
          this.setState({ handBookCompositeMetadata: topicData, siblingMetadata: topicData, pageContentType: this.contentTypeDetail.name });
        }
      }
    } else if ((this.props.webPartView === 'Chapter') && (this.props.selectedList !== undefined) && (this.props.selectedSibling !== 'All')) {
      let topicData: any;
      if (this.props.selectedChapter === 'All') {
        topicData = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), +this.props.selectedSibling, null, this.props.context);
      } else {
        topicData = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), +this.props.selectedChapter, null, this.props.context);
      }
      this.handBookCompositeMetadata = topicData;
      if (topicData[0].childNodes === undefined) {
        this.setState({ handBookCompositeMetadata: topicData, statusText: 'No item have been published for this Content Type' });
      } else {
        this.setState({ handBookCompositeMetadata: topicData });
      }
    }
    if (this.props.selectedList !== undefined) {
      let oHandbookColumn = new HandbookColumn(this.props.context);
      let scopeChoices: IDropdownOption[] = [];
      oHandbookColumn.getScopeChoices(this.props.selectedList).then((pageScopes) => {
        pageScopes[0].Choices.map((scope) => {
          scopeChoices.push({ key: scope, text: scope });
        });
      }).catch((error) => {
        console.log(error);
      });
      this.setState({ options: scopeChoices });
    }
  }

  public async componentDidUpdate(prevProps: IEhsHandbookListingProps) {
    /* Get updated data when the Web Part View is updated from the Web Part properties*/
    if (prevProps.sortByAscOrDesc !== this.props.sortByAscOrDesc) {
      this.setState({ sortByType: this.props.sortByAscOrDesc });
    }
    if (prevProps.sortByColumnName !== this.props.sortByColumnName) {
      this.setState({ sortByField: this.props.sortByColumnName });
    }
    let compositeNode = new HandbookComposite(this.props.context);
    if (((prevProps.webPartView !== this.props.webPartView) && (this.props.selectedList !== undefined))) {

      if (this.props.webPartView === strings.siblingListingOptionKey) {

        let topicData: any = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), this.parentID, null, this.props.context);
        this.handBookCompositeMetadata = topicData;
        if (topicData.length === 0) {
          this.setState({ handBookCompositeMetadata: topicData, statusText: 'No Chapters or Topics have been published' });
        } else {
          if (this.parentID !== null && topicData.length > 0) {
            this.setState({ handBookCompositeMetadata: topicData, siblingMetadata: topicData[0].childNodes || [], pageContentType: this.contentTypeDetail.name });
          } else {
            this.setState({ handBookCompositeMetadata: topicData, siblingMetadata: topicData, pageContentType: this.contentTypeDetail.name });
          }
        }

        this.setState({ selectedKey: strings.siblingListingOptionKey });
      } else if ((this.props.webPartView === 'Chapter') && (this.props.selectedSibling !== 'All')) {
        let topicData: any;
        if (this.props.selectedChapter === 'All') {
          topicData = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), +this.props.selectedSibling, null, this.props.context);
        } else {
          topicData = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), +this.props.selectedChapter, null, this.props.context);
        }
        this.handBookCompositeMetadata = topicData;
        if (topicData[0].childNodes === 0) {
          this.setState({ handBookCompositeMetadata: topicData, statusText: 'No Chapters or Topics have been published' });
        } else {
          this.setState({ handBookCompositeMetadata: topicData });
        }

        this.setState({ selectedKey: 'ChapterIndex' });
      }
    }
    /* Get updated Chapters and Topics when the selected sibling is updated from the Web Part properties*/
    if ((prevProps.selectedSibling !== this.props.selectedSibling) || (prevProps.selectedChapter !== this.props.selectedChapter)) {
      let thisSiblings = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), +this.props.selectedChapter, null, this.props.context);
      this.handBookCompositeMetadata = thisSiblings;
      if (thisSiblings[0].childNodes.length === 0) {
        this.setState({ handBookCompositeMetadata: thisSiblings, statusText: 'No Chapters or Topics have been published for this Content type' });
      } else {
        this.setState({ handBookCompositeMetadata: thisSiblings });
      }
    }
    /* Updates filter state when the filter options are updated in web part properties */
    if (JSON.stringify(prevProps.filters) !== JSON.stringify(this.props.filters)) {
      let filterState = [];
      let initialFilterState = [];
      this.props.filters.map((filter) => {
        let field = this.props.allTermSetFields.filter(termSetField => termSetField.TermSetId === filter)[0];
        if (field !== undefined) {
          filterState.push({ name: field.Title, internalName: field.InternalName, id: field.TermSetId, value: [] });
          initialFilterState.push({ name: field.Title, internalName: field.InternalName, id: field.TermSetId, value: [] });
        }
      });
      this.initialFilterState = initialFilterState;
      this.setState({ filterState, filterDisplay: this.props.filters.length === 0 ? 'none' : this.state.selectedKey === 'Siblings' ? 'none' : 'block' });
    }
    /* Gets the initial data for Sibling view when list is selected*/
    if ((this.props.webPartView === strings.siblingListingOptionKey) && (prevProps.selectedList !== this.props.selectedList) && (this.props.selectedList !== undefined)) {
      let siblingDetails = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), null, 1, this.props.context);
      this.setState({ siblingMetadata: siblingDetails });
      let topicData = await compositeNode.getChildNodes(this.props.context.pageContext.list.id.toString(), null, null, this.props.context);
      this.handBookCompositeMetadata = topicData;
      if (topicData.length === 0) {
        this.setState({ handBookCompositeMetadata: topicData, statusText: 'No Chapters or Topics have been published' });
      } else {
        this.setState({ handBookCompositeMetadata: topicData });
      }
      this.setState({ selectedKey: 'Siblings' });
    }
    if (prevProps.selectedList !== this.props.selectedList) {
      let oHandbookColumn = new HandbookColumn(this.props.context);
      let scopeChoices: IDropdownOption[] = [];
      oHandbookColumn.getScopeChoices(this.props.selectedList).then((pageScopes) => {
        pageScopes[0].Choices.map((scope) => {
          scopeChoices.push({ key: scope, text: scope });
        });
      }).catch((error) => {
        console.log(error);
      });
      this.setState({ options: scopeChoices });
    }
  }

  /**
   * Update state values as per the Pivot item selected
   */

  private handleClick(ev?: React.MouseEvent<HTMLButtonElement>, item?: IContextualMenuItem): void {
    if (item.key === 'Asc' || item.key === 'Desc') {
      this.setState({ sortByType: item.key });
    } else {
      this.setState({ sortByField: item.key });
    }
  }
  private handleLinkClick = (item: PivotItem): void => {
    if (item.props.itemKey === this.state.selectedKey && item.props.itemKey === 'AtoZ') {
      this.setState({
        sortAsc: this.state.sortAsc ? false : true,
        sortText: this.state.sortText === 'A to Z' ? 'Z to A' : 'A to Z',
        sortIcon: this.state.sortIcon === 'SortUp' ? 'SortDown' : 'SortUp'
      });
    }
    item.props.itemKey === strings.siblingListingOptionKey ? this.setState({ filterDisplay: 'none' }) : this.props.filters.length === 0 ? this.setState({ filterDisplay: 'none' }) : this.setState({ filterDisplay: 'block' });

    this.setState({
      selectedKey: item.props.itemKey
    });
  }

  /**
   * Filter topics on basis of Filter Panel's selected properties
   */
  private applyFilter(): void {
    try {
      let filteredTopics: HandbookComposite[] = [];
      let st: any;

      this.handBookCompositeMetadata.map((sub) => {
        st = { ...sub };
        st.childNodes = [];
        if (sub.hasOwnProperty('childNodes')) {
          sub.childNodes.map((chap: HandbookComposite) => {
            let ct: any = { ...chap };
            ct.childNodes = [];
            if (chap.hasOwnProperty('childNodes')) {
              chap.childNodes.map((topic: HandbookLeaf) => {
                if (this.state.filterState.every((filter) => {
                  return filter.value.every((filterValue) => {
                    let topicColumnData = topic.contentType.columns;
                    let taxonomyFieldTobeFiltered = topicColumnData.filter(termSet => (termSet.internalName === filter.internalName))[0];
                    if (taxonomyFieldTobeFiltered) {
                      if (taxonomyFieldTobeFiltered.value) {
                        if (taxonomyFieldTobeFiltered.value.length) {
                          return taxonomyFieldTobeFiltered.value.map(termValue => termValue.TermGuid).indexOf(filterValue.key) >= 0;
                        } else {
                          if (taxonomyFieldTobeFiltered.value.TermGuid) {
                            return taxonomyFieldTobeFiltered.value.TermGuid.indexOf(filterValue.key) >= 0;
                          }
                        }
                      } else {
                        return false;
                      }
                    }
                  });
                }) === true) {
                  ct.childNodes.push(topic);
                }
              });
            }

            if (ct.childNodes.length !== 0) {
              st.childNodes.push(ct);
            }
          });
        }

        if (st.childNodes.length !== 0) {
          filteredTopics.push(st);
        }
      });

      if (filteredTopics.length === 0) {
        this.setState({ showPanel: false, handBookCompositeMetadata: filteredTopics, statusText: 'No topics match the selected filtered criteria' });
      } else {
        this.setState({ showPanel: false, handBookCompositeMetadata: filteredTopics, statusText: '' });
      }
    } catch (error) {
      console.log('HandbookListing.applyFilter : ' + error);
    }
  }

  /**
   * Footer content for Filter Panel
   */
  private onRenderFooterContent = (): JSX.Element => {
    if (this.state.filterState.length === 0) {
      return (
        <div></div>
      );
    } else {
      return (
        <div>
          <PrimaryButton style={{ marginRight: '8px' }} onClick={this.applyFilter}>
            Apply
          </PrimaryButton>
          <DefaultButton onClick={() => { this.setState({ showPanel: false, handBookCompositeMetadata: this.handBookCompositeMetadata, statusText: '', filterState: this.initialFilterState }); }}>
            Clear All
          </DefaultButton>
        </div>
      );
    }
  }

  /**
   * Opens the web part property pane when Configure button of Placeholder control is clickec in page's Edit Mode
   */
  private onConfigure(): void {
    this.props.context.propertyPane.open();
  }

  // tslint:disable-next-line: max-func-body-length
  public render(): React.ReactElement<IEhsHandbookListingProps> {
    let sortOrderObj: any[] = [{
      key: strings.ascendingOptionKey,
      name: strings.ascendingOptionText,
      onClick: this.handleClick,
      canCheck: true,
      isChecked: this.state.sortByType === strings.ascendingOptionKey ? this.state.selectedKey === 'AtoZ' ? false : true : false,
      iconProps: {
        iconName: strings.ascendingOptionText
      },
      disabled: this.state.selectedKey === 'AtoZ' ? true : false
    },
    {
      key: strings.descendingOptionKey,
      name: strings.descendingOptionText,
      onClick: this.handleClick,
      canCheck: true,
      isChecked: this.state.sortByType === strings.descendingOptionKey ? this.state.selectedKey === 'AtoZ' ? false : true : false,
      iconProps: {
        iconName: strings.descendingOptionText
      },
      disabled: this.state.selectedKey === 'AtoZ' ? true : false
    },
    {
      key: null,
      itemType: ContextualMenuItemType.Divider
    },
    {
      key: strings.sortByDefaultColKey,
      name: strings.sortByDefaultColText,
      onClick: this.handleClick,
      canCheck: true,
      isChecked: this.state.sortByField === strings.sortByDefaultColKey ? true : false,
      iconProps: {
        style: { display: 'none' }
      }
    }];

    if (this.props.additionalField.length > 0) {
      this.props.additionalField.map((value: string) => {
        if (value !== strings.sortByDefaultColKey) {
          sortOrderObj.push(
            {
              key: value,
              name: value,
              onClick: this.handleClick,
              canCheck: true,
              checked: this.state.sortByField === value ? true : false,
              iconProps: {
                style: { display: 'none' }
              }
            }
          );
        }
      });
    }

    if (this.props.configured) {
      return (
        <div className={EhsHandbookListingModuleScss.ehsHandbookListing}>
          <div className={EhsHandbookListingModuleScss.container}>
            <div className={EhsHandbookListingModuleScss.row}>
              <div className='ms-Grid-col ms-u-md9 ms-u-sm9 ms-u-xs9'>
                {this.props.webPartView === strings.siblingListingOptionKey &&
                  <Pivot headersOnly={true}
                    selectedKey={this.state.selectedKey}
                    onLinkClick={this.handleLinkClick}
                  >
                    <PivotItem headerText={this.state.pageContentType} itemKey={strings.siblingListingOptionKey} />
                    <PivotItem headerText='Index' itemKey='EHSIndex' />
                  </Pivot>
                }
                {this.props.webPartView === 'Chapter' &&
                  <Pivot headersOnly={true}
                    selectedKey={this.state.selectedKey}
                    onLinkClick={this.handleLinkClick}
                  >
                    <PivotItem headerText='Index' itemKey='ChapterIndex' />
                    <PivotItem headerText={this.state.sortText} itemIcon={this.state.sortIcon} itemKey='AtoZ' />
                  </Pivot>
                }
              </div>
              <div className='ms-Grid-col ms-u-md3 ms-u-sm3 ms-u-xs3'>
                <span style={{ display: this.state.filterDisplay }}><Icon iconName='Filter' styles={{ root: { float: 'right', fontSize: 'larger', margin: '10px' } }} onClick={() => { this.setState({ showPanel: true }); }} /></span>
                <DefaultButton
                  text='Sort By'
                  id='ContextualMenuButton'
                  iconProps={{ iconName: 'SortLines' }}
                  menuIconProps={{ iconName: '' }}
                  menuProps={{
                    shouldFocusOnMount: true,
                    items: sortOrderObj
                  }}
                />
              </div>
            </div>
            <div className={EhsHandbookListingModuleScss.row}>
              <div>
                {this.state.selectedKey === strings.siblingListingOptionKey &&
                  <Subjects siblingMetadata={this.state.siblingMetadata}
                    context={this.props.context}
                    logsSiteUrl={this.props.logsSiteUrl}
                    logsTitle={this.props.logsTitle}
                    writeToDebug={this.props.writeToDebug}
                    sortOrder={this.props.sortByColumnName}
                    sortType={this.props.sortByAscOrDesc}
                    sortByType={this.state.sortByType}
                    sortByField={this.state.sortByField}
                    additionalColumn={this.props.additionalField}
                  />
                }
                {this.state.selectedKey === 'EHSIndex' &&
                  <EHSIndex handBookCompositeMetadata={this.state.handBookCompositeMetadata}
                    context={this.props.context}
                    selectedList={this.props.selectedList}
                    options={this.state.options}
                    includeChapterLinks={this.props.includeChapterLinks}
                    logsSiteUrl={this.props.logsSiteUrl}
                    logsTitle={this.props.logsTitle}
                    writeToDebug={this.props.writeToDebug}
                    sortOrder={this.props.sortByColumnName}
                    sortType={this.props.sortByAscOrDesc}
                    sortByType={this.state.sortByType}
                    sortByField={this.state.sortByField}
                    additionalColumn={this.props.additionalField}
                  />
                }
                {this.state.selectedKey === 'ChapterIndex' &&
                  <ChapterIndex isContributor={this.props.isContributor}
                    isOwner={this.props.isOwner}
                    context={this.props.context}
                    selectedList={this.props.selectedList}
                    options={this.state.options}
                    templateUrl={this.props.topicTemplateUrl}
                    includeChapterLinks={this.props.includeChapterLinks}
                    handBookCompositeMetadata={this.state.handBookCompositeMetadata}
                    logsSiteUrl={this.props.logsSiteUrl}
                    logsTitle={this.props.logsTitle}
                    writeToDebug={this.props.writeToDebug}
                    selectetItem={this.props.selectedChapter}
                    sortOrder={this.props.sortByColumnName}
                    sortType={this.props.sortByAscOrDesc}
                    sortByType={this.state.sortByType}
                    sortByField={this.state.sortByField}
                    additionalColumn={this.props.additionalField}
                  />
                }
                {this.state.selectedKey === 'AtoZ' &&
                  <AtoZ sortAsc={this.state.sortAsc}
                    handBookCompositeMetadata={this.state.handBookCompositeMetadata}
                    context={this.props.context}
                    selectedList={this.props.selectedList}
                    options={this.state.options}
                    logsSiteUrl={this.props.logsSiteUrl}
                    logsTitle={this.props.logsTitle}
                    writeToDebug={this.props.writeToDebug}
                    additionalColumn={this.props.additionalField}
                    selectetItem={this.props.selectedChapter}
                    sortByField={this.state.sortByField}
                  />
                }
                <p style={{ padding: '15px' }}>{this.state.statusText}</p>
              </div>
            </div>
            <Panel isOpen={this.state.showPanel}
              headerText='Filter'
              onDismiss={() => this.setState({ showPanel: false })}
              onRenderFooterContent={this.onRenderFooterContent}
            >
              {this.state.filterState.length === 0 &&
                <p>Please configure filters in the web part properties for use</p>
              }
              {this.state.filterState.length > 0 &&
                this.state.filterState.map((filter) => {
                  return (
                    <TaxonomyPicker
                      allowMultipleSelections={filter.allowMultipleValues}
                      termsetNameOrID={filter.termsetId}
                      anchorId={filter.id}
                      panelTitle={`Select ${filter.name}`}
                      label={filter.name}
                      context={this.props.context}
                      isTermSetSelectable={false}
                      initialValues={filter.value}
                      onChange={(terms: IPickerTerms) => {
                        let filters = [...this.state.filterState];
                        filters.filter(f => f.id === filter.id)[0].value = terms;
                        this.setState({ filterState: filters });
                      }}
                    />
                  );
                })}
            </Panel>
          </div>
        </div>
      );
    } else {
      return (
        <Placeholder iconName='Edit'
          iconText='Configure your web part'
          description='Please configure the web part.'
          buttonLabel='Configure'
          hideButton={this.props.displayMode === DisplayMode.Read}
          onConfigure={this.onConfigure} />
      );
    }
  }
}