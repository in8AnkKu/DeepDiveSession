import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import * as strings from 'EhsHandbookPageDetailsWebPartStrings';
import EhsHandbookPageDetails from './components/EhsHandbookPageDetails';
import { IEhsHandbookPageDetailsProps } from './components/IEhsHandbookPageDetailsProps';
import { HandbookColumn } from '../../BAL/HandbookColumn';

export interface IEhsHandbookPageDetailsWebPartProps {
  layoutType: string;
  logoUrl: string;
  pageProperties: string[];
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
}

export default class EhsHandbookPageDetailsWebPart extends BaseClientSideWebPart<IEhsHandbookPageDetailsWebPartProps> {

  private allPageProperties: IPropertyPaneDropdownOption[];

  public render(): void {
    const element: React.ReactElement<IEhsHandbookPageDetailsProps> = React.createElement(
      EhsHandbookPageDetails,
      {
        context: this.context,
        layoutType: this.properties.layoutType,
        logoUrl: this.properties.logoUrl,
        pageProperties: this.properties.pageProperties,
        configured: this.properties.pageProperties === undefined ? false : true,
        displayMode: this.displayMode,
        logsSiteUrl: this.properties.logsSiteUrl,
        logsTitle: escape(this.properties.logsTitle),
        writeToDebug: this.properties.writeToDebug
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(async _ => {
      if (this.properties.logoUrl === undefined) {
        this.properties.logoUrl = this.context.pageContext.web.absoluteUrl + '/SiteAssets/sitelogo.png';
      }
      if (this.properties.layoutType === undefined) {
        this.properties.layoutType = 'basic';
      }
      let pageContentTypeFields = await (new HandbookColumn(this.context)).getAllColumnsForContentType(this.context.pageContext.list.id.toString(), this.context.pageContext.listItem.id);
      let fieldsProps: IPropertyPaneDropdownOption[] = [];
      let dependentLookupInternalNames: any = [];
      pageContentTypeFields.map((field) => {
        if (field.DependentLookupInternalNames) {
          field.DependentLookupInternalNames.map((value) => {
            dependentLookupInternalNames.push(value);
          });
        }
      });
      pageContentTypeFields.map((termSetField) => {
        if (dependentLookupInternalNames.indexOf(termSetField.InternalName) === -1) {
          fieldsProps.push({ key: termSetField.InternalName, text: termSetField.Title });
        }
      });
      fieldsProps.push({ key: 'OData__UIVersionString', text: 'Version' });
      this.allPageProperties = fieldsProps;
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
                PropertyFieldMultiSelect('pageProperties', {
                  key: 'pageProperties',
                  label: 'Select the page properties to be displayed',
                  options: this.allPageProperties,
                  selectedKeys: this.properties.pageProperties ? this.properties.pageProperties : []
                }),
                PropertyPaneChoiceGroup('layoutType', {
                  label: 'Select layout',
                  options: [
                    {
                      key: 'basic',
                      text: 'Basic Layout',
                      checked: true
                    },
                    {
                      key: 'advanced',
                      text: 'Advanced Layout'
                    }
                  ]
                }),
                PropertyPaneTextField('logoUrl', {
                  label: 'Image URL',
                  value: escape(this.properties.logoUrl),
                  disabled: escape(this.properties.layoutType) === 'basic' ? true : false
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