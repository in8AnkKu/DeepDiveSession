import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import * as strings from 'EhsHandbookAddNewPageWebPartStrings';
import EhsHandbookAddNewPage from './components/EhsHandbookAddNewPage';
import { IEhsHandbookAddNewPageProps } from './components/IEhsHandbookAddNewPageProps';
import { HandbookColumn } from '../../BAL/HandbookColumn';
import { HandbookComposite } from '../../BAL/HandbookComposite';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

export interface IEhsHandbookAddNewPageWebPartProps {
  pageType: string;
  lists: string;
  templateUrl: string;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
}

export default class EhsHandbookAddNewPageWebPart extends BaseClientSideWebPart<IEhsHandbookAddNewPageWebPartProps> {
  /* List of all Content Types for the selected List*/
  private allContentTypesList: IPropertyPaneDropdownOption[];
  /* Scope of the parent page on which the web part is added*/
  private parentPageScope: string = '';
  private pageScopes: IDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<IEhsHandbookAddNewPageProps> = React.createElement(
      EhsHandbookAddNewPage,
      {
        pageType: escape(this.properties.pageType),
        context: this.context,
        selectedList: this.properties.lists,
        templateUrl: this.properties.templateUrl,
        configured: (this.properties.lists && this.properties.pageType && this.properties.templateUrl) ? true : false,
        displayMode: this.displayMode,
        parentPageScope: this.parentPageScope,
        logsSiteUrl: this.properties.logsSiteUrl,
        logsTitle: escape(this.properties.logsTitle),
        writeToDebug: this.properties.writeToDebug,
        pageScopes: this.pageScopes
      }
    );

    /* Render the Add New Page Web Part on the basis of user permissions*/
    if (this.properties.lists) {
      let handbookComposite = new HandbookComposite(this.context);
      handbookComposite.checkUserPermissions(escape(this.properties.lists), 'ManageWeb').then((perms) => {
      try {
        if (perms === true) {
          ReactDom.render(element, this.domElement);
        }
      } catch (error) {
        console.log(error);
      }
      });
    }
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(async _ => {
      if (this.properties.lists) {
        let handbookComposite = new HandbookComposite(this.context);
        let parentPage = await handbookComposite.getListItem(escape(this.properties.lists), this.context.pageContext.listItem.id);
        this.parentPageScope = parentPage.Scope;
        let handbookColumnContext = new HandbookColumn(this.context);
        let scopeChoices: IDropdownOption[] = [];
        let pageScopes = await handbookColumnContext.getScopeChoices(escape(this.properties.lists));
        pageScopes[0].Choices.map((scope) => {
          scopeChoices.push({ key: scope, text: scope });
        });
        this.pageScopes = scopeChoices;
      }
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
      let handbookComposite = new HandbookComposite(this.context);
      handbookComposite.loadContentTypes(escape(this.properties.lists)).then((allContentTypesItems) => {
        try {
          this.allContentTypesList = allContentTypesItems;
          this.context.propertyPane.refresh();
        } catch (error) {
          console.log(error);
        }
      });
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string) {
    /* Get updated list of all content types corresponding to the selected list*/
    if (propertyPath === 'lists') {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
      let handbookComposite = new HandbookComposite(this.context);
      handbookComposite.loadContentTypes(escape(this.properties.lists)).then((allSubjectItems) => {
        try {
          this.allContentTypesList = allSubjectItems;
          this.context.propertyPane.refresh();
        } catch (error) {
          console.log(error);
        }
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
                PropertyPaneDropdown('pageType', {
                  label: 'Page Content Type',
                  options: this.allContentTypesList
                }),
                PropertyPaneTextField('templateUrl', {
                  label: 'Template URL'
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