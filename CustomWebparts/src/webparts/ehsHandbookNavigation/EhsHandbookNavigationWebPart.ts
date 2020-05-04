import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import * as strings from 'EhsHandbookNavigationWebPartStrings';
import EhsHandbookNavigation from './components/EhsHandbookNavigation';
import { IEhsHandbookNavigationProps } from './components/IEhsHandbookNavigationProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

export interface IEhsHandbookNavigationWebPartProps {
  selectedSubject: string;
  lists: string;
  includeChapterLinks: boolean;
  showSubjectLink: boolean;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
}

export default class EhsHandbookNavigationWebPart extends BaseClientSideWebPart<IEhsHandbookNavigationWebPartProps> {
  //List of all Subjects in the Handbook
  private allSubjectsList: IPropertyPaneDropdownOption[];

  protected onInit(): Promise<void> {
    if (this.properties.showSubjectLink === undefined) {
      this.properties.showSubjectLink = true;
    }
    if (this.properties.includeChapterLinks === undefined) {
      this.properties.includeChapterLinks = true;
    }
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IEhsHandbookNavigationProps> = React.createElement(
      EhsHandbookNavigation,
      {
        context: this.context,
        selectedList: this.properties.lists,
        configured: (this.properties.lists) ? true : false,
        displayMode: this.displayMode,
        logsSiteUrl: this.properties.logsSiteUrl,
        logsTitle: escape(this.properties.logsTitle),
        writeToDebug: this.properties.writeToDebug
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string) {
    /**
     * Load updated subjects when the list is selected
     */
    if (propertyPath === 'lists') {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
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
