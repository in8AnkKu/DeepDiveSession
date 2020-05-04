import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { HandbookComposite } from '../../BAL/HandbookComposite';
import * as strings from 'EhsHandbookListingWebPartStrings';
import EhsHandbookListing from './components/EhsHandbookListing';
import { IEhsHandbookListingProps } from './components/IEhsHandbookListingProps';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { HandbookContentType } from '../../BAL/HandbookContentType';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

export interface IEhsHandbookListingWebPartProps {
  webPartView: string;
  selectedSibling: string;
  selectedChapter: string;
  lists: string;
  topicTemplateUrl: string;
  includeChapterLinks: boolean;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
  filters: string[];
  sortByColumnName: string;
  sortByAscOrDesc: string;
  additionalField: string[];
  defaultShowChildCount: number;
  showChildOrGrandChild: string;
}

export default class EhsHandbookListingWebPart extends BaseClientSideWebPart<IEhsHandbookListingWebPartProps> {
  //List of all sibling in the Handbook
  private allSubjectsList: IPropertyPaneDropdownOption[] = [];
  //List of all Chapters in the selected Handbook Subject
  private allChaptersList: IPropertyPaneDropdownOption[] = [];
  private allFilters: IPropertyPaneDropdownOption[] = [];
  private allTermSetFields: any[] = [];
  private isContributor: boolean;
  private isOwner: boolean;
  private additionFieldOptions: IPropertyPaneDropdownOption[];
  private ddlSortByColOptions: IPropertyPaneDropdownOption[];
  private oHandbookComposite: HandbookComposite;

  public render(): void {
    const element: React.ReactElement<IEhsHandbookListingProps> = React.createElement(
      EhsHandbookListing,
      {
        webPartView: escape(this.properties.webPartView),
        selectedSibling: escape(this.properties.selectedSibling),
        selectedChapter: escape(this.properties.selectedChapter),
        context: this.context,
        selectedList: this.properties.lists,
        topicTemplateUrl: this.properties.topicTemplateUrl,
        isContributor: this.isContributor,
        isOwner: this.isOwner,
        configured: (this.properties.lists && (escape(this.properties.webPartView) === strings.childListingOptionKey ? (this.properties.selectedSibling && this.properties.topicTemplateUrl) : true)) ? true : false,
        displayMode: this.displayMode,
        includeChapterLinks: this.properties.includeChapterLinks,
        logsSiteUrl: this.properties.logsSiteUrl,
        logsTitle: escape(this.properties.logsTitle),
        writeToDebug: this.properties.writeToDebug,
        filters: this.properties.filters,
        allTermSetFields: this.allTermSetFields,
        sortByColumnName: escape(this.properties.sortByColumnName),
        sortByAscOrDesc: escape(this.properties.sortByAscOrDesc),
        additionalField: this.properties.additionalField,
        showChildOrGrandChild: escape(this.properties.showChildOrGrandChild),
        defaultShowChildCount: this.properties.defaultShowChildCount !== undefined ? this.properties.defaultShowChildCount : strings.defaultShowChildCount
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(async _ => {
      this.ddlSortByColOptions = [{
        key: strings.sortByDefaultColKey,
        text: strings.sortByDefaultColText
      }];

      if (this.properties.selectedChapter === undefined) {
        this.properties.selectedChapter = 'All';
      }
      if (this.properties.filters === undefined) {
        this.properties.filters = [];
      }
      if (this.properties.includeChapterLinks === undefined) {
        this.properties.includeChapterLinks = true;
      }
      if (this.properties.sortByAscOrDesc === undefined) {
        this.properties.sortByAscOrDesc = strings.ascendingOptionKey;
      }
      if (this.properties.sortByColumnName === undefined) {
        this.properties.sortByColumnName = strings.sortByDefaultColKey;
      }
      if (this.properties.additionalField === undefined) {
        this.properties.additionalField = [];
      }
      if (this.properties.lists !== undefined) {
        //Check current user's permissions to render the Add Topic link in ChapterIndex component
        const handbookComposite = new HandbookComposite(this.context);
        let isContributor = await handbookComposite.checkUserPermissions(escape(this.properties.lists), 'EditListItems');
        this.isContributor = isContributor;
        let isOwner = await handbookComposite.checkUserPermissions(escape(this.properties.lists), 'ManageWeb');
        this.isOwner = isOwner;
        let allTermSetFields = await handbookComposite.loadTermsetFields(escape(this.properties.lists), ['Title', 'TermSetId', 'InternalName', 'AnchorId', 'AllowMultipleValues']);
        this.allTermSetFields = allTermSetFields;
        if (this.properties.filters === undefined) {
          let filters: string[] = [];
          allTermSetFields.map((termSetField) => {
            filters.push(termSetField.AnchorId);
          });
          this.properties.filters = filters;
        }
      }
    }).catch((error) => {
      console.log(error);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart() {
    if (this.properties.lists !== undefined) {
      this.getAdditionalColumns();
      this.fillSortingOptions([]);
    }

    if (this.properties.lists !== undefined) {
      (new HandbookContentType(this.context)).populateContentTypeLevel(this.properties.lists).then((allContentTypes) => {
        let handbookComposite = new HandbookComposite(this.context);
        handbookComposite.loadRootLevelData(this.properties.lists, allContentTypes[0].contentTypeID).then((allSubjectItems) => {
          let sortedallSubjectItems: any = allSubjectItems.sort((item1, item2) => item1.Title.localeCompare(item2.Title, 'en-u-kn-true'));
          this.allSubjectsList = [];
          sortedallSubjectItems.map((currentSubject: any) => {
            this.allSubjectsList.push({ key: currentSubject.PageId, text: currentSubject.Title });
          });
          this.context.propertyPane.refresh();
        }).catch((error) => {
          console.log(error);
        });
      }).catch((error) => {
        console.log(error);
      });

      let ohandbookComposite = new HandbookComposite(this.context);
      ohandbookComposite.loadTermsetFields(escape(this.properties.lists), ['Title', 'TermSetId', 'AnchorId', 'AllowMultipleValues']).then((termSetFields) => {
        let allFilters: IPropertyPaneDropdownOption[] = [];
        termSetFields.map((termSetField) => {
          let fiterKey = termSetField.AnchorId === '00000000-0000-0000-0000-000000000000' ? termSetField.TermSetId + 'TermSet' : termSetField.AnchorId;
          allFilters.push({ key: fiterKey, text: termSetField.Title });
        });
        this.allFilters = allFilters;
        this.context.propertyPane.refresh();
      }).catch((error) => {
        console.log(error);
      });
    }
    if (escape(this.properties.selectedSibling) !== 'All') {
      this.allChaptersList = [];
      let handbookComposite = new HandbookComposite(this.context);
      handbookComposite.getChildNodes(escape(this.properties.lists), +this.properties.selectedSibling, 2, this.context).then((allChapterItems) => {

        this.allChaptersList.push({ key: 'All', text: 'All' });
        if (allChapterItems[0] !== undefined) {
          let sortedAllChapterItems: any = allChapterItems[0].childNodes.sort((item1, item2) => item1.title.localeCompare(item2.title, 'en-u-kn-true'));
          sortedAllChapterItems.map((currentChapter: any) => {
            this.allChaptersList.push({ key: currentChapter.Id, text: currentChapter.Title });
          });
        }
        this.context.propertyPane.refresh();
      }).catch((error) => {
        console.log(error);
      });
    }
  }

  /**
   * function to fill the drop down options of sorting property
   * @param oldValue what is the old value of the property
   */
  private fillSortingOptions(oldValue: string[]): void {
    try {
      if (this.properties.additionalField.length > strings.maxCountInAddCol) {
        alert(strings.warningMessageofAdditionalFieldSelection);
        this.properties.additionalField = oldValue;
      } else {
        this.ddlSortByColOptions = [];
        this.ddlSortByColOptions = [{
          key: strings.sortByDefaultColKey,
          text: strings.sortByDefaultColText
        }];
        this.properties.additionalField.map((colName) => {
          if (colName !== strings.sortByDefaultColKey) {
            this.ddlSortByColOptions.push({
              key: colName,
              text: colName
            });
          }
        });
      }
    } catch (error) {
      console.log(error);
    }
  }

  /**
   * function to get the additionalcolumns in property pane
   */
  private async getAdditionalColumns(): Promise<void> {
    try {
      let oHandbookContentType = new HandbookContentType(this.context);
      let getAllColumn = await oHandbookContentType.getAllColumnsForContentType(this.properties.lists, this.context.pageContext.listItem.id);
      this.additionFieldOptions = [];
      getAllColumn.map((column) => {
        this.additionFieldOptions.push({
          key: column.name,
          text: column.name
        });
      });
      this.context.propertyPane.refresh();
    } catch (error) {
      console.log(error);
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string) {
    /* Get updated list of all content types corresponding to the selected list*/
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    switch (propertyPath) {
      case `lists`:
        this.context.propertyPane.refresh();
        this.getAdditionalColumns();
        break;
      case `additionalField`:
        this.context.propertyPane.refresh();
        this.fillSortingOptions(<string[]><unknown>oldValue);
        break;
    }
    /**
     * Load updated Siblings when the list is selected
     */
    if (propertyPath === 'lists') {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();

      if (this.properties.lists !== undefined) {
        (new HandbookContentType(this.context)).populateContentTypeLevel(this.properties.lists).then((allContentTypes) => {
          let handbookComposite = new HandbookComposite(this.context);
          handbookComposite.loadRootLevelData(this.properties.lists, allContentTypes[0].contentTypeID).then((allSubjectItems) => {
            let sortedallSubjectItems: any = allSubjectItems.sort((item1, item2) => item1.Title.localeCompare(item2.Title, 'en-u-kn-true'));
            this.allSubjectsList = [];
            sortedallSubjectItems.map((currentSubject: any) => {
              this.allSubjectsList.push({ key: currentSubject.PageId, text: currentSubject.Title });
            });
            this.context.propertyPane.refresh();
          }).catch((error) => {
            console.log(error);
          });
        }).catch((error) => {
          console.log(error);
        });

        let ohandbookComposite = new HandbookComposite(this.context);
        ohandbookComposite.loadTermsetFields(escape(this.properties.lists), ['Title', 'TermSetId', 'InternalName', 'AnchorId', 'AllowMultipleValues']).then((termSetFields) => {
          this.allTermSetFields = termSetFields;
          let allFilters: IPropertyPaneDropdownOption[] = [];
          termSetFields.map((termSetField) => {
            let fiterKey = termSetField.AnchorId === '00000000-0000-0000-0000-000000000000' ? termSetField.TermSetId + 'TermSet' : termSetField.AnchorId;
            allFilters.push({ key: fiterKey, text: termSetField.Title });
          });
          this.allFilters = allFilters;
          this.context.propertyPane.refresh();
        }).catch((error) => {
          console.log(error);
        });
      }
    } else if (propertyPath === 'selectedSibling') {
      this.allChaptersList = [];
      let ohandbookComposite = new HandbookComposite(this.context);
      ohandbookComposite.getChildNodes(escape(this.properties.lists), +this.properties.selectedSibling, 2, this.context).then((allChapterItems) => {
        this.allChaptersList.push({ key: 'All', text: 'All' });

        if (allChapterItems[0].childNodes !== undefined) {
          let sortedAllChapterItems: any = allChapterItems[0].childNodes.sort((item1, item2) => item1.title.localeCompare(item2.title, 'en-u-kn-true'));
          sortedAllChapterItems.map((currentChapter: any) => {
            this.allChaptersList.push({ key: currentChapter.id, text: currentChapter.title });
          });
        }
        this.context.propertyPane.refresh();
      }).catch((error) => {
        console.log(error);
      });
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: escape(this.properties.lists),
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneChoiceGroup('webPartView', {
                  label: strings.webPartViewLabel,
                  options: [
                    { key: strings.siblingListingOptionKey, text: strings.siblingListingOptionText },
                    { key: strings.childListingOptionKey, text: strings.childListingOptionText }
                  ]
                }),
                PropertyPaneChoiceGroup('showChildOrGrandChild', {
                  label: strings.showChildOrGrandChildLabel,
                  options: [
                    {
                      key: strings.childOptionKey,
                      text: strings.childOptionText
                    },
                    {
                      key: strings.childGrandChildOptionKey,
                      text: strings.childGrandChildOptionText,
                      checked: true
                    }
                  ]
                }),
                PropertyFieldNumber(`defaultShowChildCount`, {
                  key: `defaultShowChildCount`,
                  label: strings.defaultShowChildCountLabel,
                  description: strings.defaultShowChildCountDesc,
                  value: this.properties.defaultShowChildCount !== undefined ? this.properties.defaultShowChildCount : strings.defaultShowChildCount,
                  minValue: 1,
                  disabled: this.properties.showChildOrGrandChild === strings.childGrandChildOptionKey ? false : true
                }),
                PropertyFieldMultiSelect('additionalField', {
                  key: 'additionalField',
                  label: strings.additionFieldLabel,
                  options: this.additionFieldOptions,
                  selectedKeys: this.properties.additionalField
                }),
                PropertyPaneDropdown(`sortByColumnName`, {
                  label: strings.sortByColumnNameLabel,
                  options: this.ddlSortByColOptions,
                  selectedKey: this.properties.sortByColumnName
                }),
                PropertyPaneChoiceGroup(`sortByAscOrDesc`, {
                  label: strings.sortByAscOrDescLabel,
                  options: [
                    {
                      key: strings.ascendingOptionKey,
                      text: strings.ascendingOptionText,
                      checked: true
                    },
                    {
                      key: strings.descendingOptionKey,
                      text: strings.descendingOptionText
                    }
                  ]
                }),
                PropertyFieldMultiSelect('filters', {
                  key: 'filters',
                  label: 'Select the filters',
                  options: this.allFilters,
                  selectedKeys: this.properties.filters ? this.properties.filters : []
                }),
                PropertyPaneToggle('includeChapterLinks', {
                  label: 'Include Chapter Links',
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneDropdown('selectedSibling', {
                  label: 'Select Sibling',
                  options: this.allSubjectsList,
                  disabled: this.properties.webPartView === 'Sibling' ? true : false
                }),
                PropertyPaneDropdown('selectedChapter', {
                  label: 'Select Chapter',
                  options: this.allChaptersList,
                  disabled: escape(this.properties.webPartView) === 'Sibling' ? true : false
                }),
                PropertyPaneTextField('topicTemplateUrl', {
                  label: 'New Topic Template URL',
                  disabled: escape(this.properties.webPartView) === 'Sibling' ? true : false
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'Error logging properties',
              groupFields: [
                PropertyPaneTextField('logsSiteUrl', {
                  label: 'Site URL for Error Logging'
                }),
                PropertyPaneTextField('logsTitle', {
                  label: 'Application Title for Error Log'
                }),
                PropertyPaneToggle('writeToDebug', {
                  label: 'Write to Debug',
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            }
          ]
        }
      ]
    };
  }

}
