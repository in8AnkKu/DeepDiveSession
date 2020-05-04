
import * as React from 'react';
import EhsHandbookPageDetailsModuleScss from './EhsHandbookPageDetails.module.scss';
import { IEhsHandbookPageDetailsProps } from './IEhsHandbookPageDetailsProps';
import { IEhsHandbookPageDetailsState } from './IEhsHandbookPageDetailsState';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { HandbookComposite } from '../../../BAL/HandbookComposite';
import { HandbookColumn } from '../../../BAL/HandbookColumn';

export default class EhsHandbookPageDetails extends React.Component<IEhsHandbookPageDetailsProps, IEhsHandbookPageDetailsState> {
  constructor(props: IEhsHandbookPageDetailsProps) {
    super(props);

    this.state = {
      pageDetails: [],
      fieldType: [],
      fieldLabel: [],
      singleTermLabel: [],
      fieldValue: []
    };
    this.getPageDetails = this.getPageDetails.bind(this);
    this.onConfigure = this.onConfigure.bind(this);

  }

  public componentDidMount() {
    this.getPageDetails();
  }

  public componentDidUpdate(prevProps: IEhsHandbookPageDetailsProps) {
    if (JSON.stringify(prevProps.pageProperties) !== JSON.stringify(this.props.pageProperties)) {
      this.getPageDetails();
    }
  }

  /**
  * Get page properties to be rendered
  */
  private async getPageDetails() {
    try {
      let literalVersion = 'OData__UIVersionString';
      let pageObject = new HandbookComposite(this.props.context);
      let fieldType = {};
      let fieldLabel = {};
      let fieldValue = {};
      let pageData = await pageObject.getPageDetails(this.props.context.pageContext.list.id.toString(), this.props.context.pageContext.listItem.id, this.props.context);
      let pageDataFromDirectAPI = await (new HandbookComposite(this.props.context)).getPageDetailsById(this.props.context.pageContext.list.id.toString(), this.props.context.pageContext.listItem.id);
      let columnData = pageData.contentType.columns;
      this.setState({ pageDetails: columnData });
      columnData.map((item) => {
        fieldType[item.internalName] = item.columnType;
        fieldLabel[item.internalName] = item.name;
        fieldValue[item.internalName] = item.value;
      });
      fieldType[literalVersion] = 'Text';
      fieldLabel[literalVersion] = 'Version';
      fieldValue[literalVersion] = pageDataFromDirectAPI.OData__UIVersionString;

      this.setState({ fieldType: fieldType, fieldLabel: fieldLabel, fieldValue: fieldValue });

      let singleTermIds: any[] = [];
      columnData.map((item) => {
        if (item.columnType === 'TaxonomyFieldType') {
          singleTermIds.push('ID eq ' + item.value.Label);
        }
      });
      let termStoreData = await (new HandbookColumn(this.props.context)).loadSingleTaxonomyValue('TaxonomyHiddenList', ['ID', 'Title'], singleTermIds.join(' OR '));
      let singleTermData: any = [];
      termStoreData.map((term) => {
        singleTermData.push({ 'Id': term.ID, 'Label': term.Title });
      });
      this.setState({ singleTermLabel: singleTermData });
    } catch (error) {
      console.log('HandbookPageDetails.getPageDetails : ' + error);
    }
  }

  /**
   * Opens the web part property pane when Configure button of Placeholder control is clickec in page's Edit Mode
   */
  private onConfigure(): void {
    this.props.context.propertyPane.open();
  }

  public render(): React.ReactElement<IEhsHandbookPageDetailsProps> {
    let literalModified = 'Modified';
    let literalCreated = 'Created';
    let literalScope = 'Scope';
    let literalVersion = 'OData__UIVersionString';
    let months: string[] = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    if (this.props.configured) {
      const fieldType = this.state.fieldType;
      const fieldValue = this.state.fieldValue;
      const fieldLabel = this.state.fieldLabel;
      const context = this.props.context;
      const absoluteUrl = context.pageContext.web.absoluteUrl;
      const selectedPageProperties: any[] = [];

      this.props.pageProperties.map(async (field) => {
        let fieldValueText: any = '---';

        if (fieldType[field] === 'Text' || fieldType[field] === 'Number' || fieldType[field] === 'Choice') {
          fieldValueText = fieldValue[field] === null ? fieldValueText : fieldValue[field];
          if ((field.indexOf('Created_x0020_By') >= 0) || (field.indexOf('Modified_x0020_By') >= 0)) {
            let fieldValueTextSplitArray = fieldValueText.split('|');
            fieldValueText = fieldValueTextSplitArray[fieldValueTextSplitArray.length - 1];
          }
          selectedPageProperties.push({ key: fieldLabel[field], text: fieldValueText, type: fieldType[field] });
        }
        if (fieldType[field] === 'File') {
          if (fieldValueText) {
            fieldValueText = fieldValue[field] === null ? fieldValueText : fieldValue[field].substr(0, fieldValue[field].lastIndexOf('.'));
          }
          selectedPageProperties.push({ key: fieldLabel[field], text: fieldValueText, type: fieldType[field] });
        }
        if (fieldType[field] === 'Boolean') {
          fieldValueText = fieldValue[field] === null ? fieldValueText : fieldValue[field] === true ? 'True' : 'False';
          selectedPageProperties.push({ key: fieldLabel[field], text: fieldValueText, type: fieldType[field] });
        }
        if (fieldType[field] === 'MultiChoice') {
          fieldValueText = fieldValue[field] === null ? fieldValueText : fieldValue[field].map(item => item).join(', ');
          selectedPageProperties.push({ key: fieldLabel[field], text: fieldValueText, type: fieldType[field] });
        }
        if (fieldType[field] === 'DateTime') {
          let dateFieldValue = fieldValue[field] === null ? fieldValueText : fieldValue[field];
          if (dateFieldValue !== fieldValueText) {
            dateFieldValue = new Date(dateFieldValue);
            let dateMonth = dateFieldValue.getMonth();
            let dateDay = dateFieldValue.getDate();
            let dateYear = dateFieldValue.getFullYear();
            dateFieldValue = dateDay + ' ' + months[dateMonth] + ' ' + dateYear;
          }
          selectedPageProperties.push({ key: fieldLabel[field], text: dateFieldValue, type: fieldType[field] });
        }
        if (fieldType[field] === 'User') {
          let userTitle = fieldValue[field] !== undefined && fieldValue[field] !== null ?
            <Link target='_blank' data-interception='off' href={absoluteUrl + '/_layouts/15/me.aspx/?p=' + fieldValue[field].EMail + '&v=work'}>{fieldValue[field].Title}</Link>
            : fieldValueText;
          selectedPageProperties.push({ key: fieldLabel[field], text: userTitle, type: fieldType[field] });
        }
        if (fieldType[field] === 'Lookup') {
          fieldValueText = fieldValue[field] !== undefined && fieldValue[field] !== null ? fieldValue[field].Title : fieldValueText;
          selectedPageProperties.push({ key: fieldLabel[field], text: fieldValueText, type: fieldType[field] });
        }
        if (fieldType[field] === 'TaxonomyFieldTypeMulti') {
          fieldValueText = fieldValue[field].length === 0 ? fieldValueText : fieldValue[field].map((term) => { return term.Label; }).join(', ');
          selectedPageProperties.push({ key: fieldLabel[field], text: fieldValueText, type: fieldType[field] });
        }
        if (fieldType[field] === 'TaxonomyFieldType') {
          if (fieldValue[field] !== null) {
            this.state.singleTermLabel.map((termid) => {
              if (fieldValue[field].Label === termid.Id.toString()) {
                fieldValueText = termid.Label;
              }
            });
          }
          selectedPageProperties.push({ key: fieldLabel[field], text: fieldValueText, type: fieldType[field] });
        }
      });

      return (
        <div className={EhsHandbookPageDetailsModuleScss.ehsHandbookPageDetails}>
          <div className={EhsHandbookPageDetailsModuleScss.container}>
            <div className={EhsHandbookPageDetailsModuleScss.row}>
              {this.props.layoutType === 'advanced' &&
                <div className={EhsHandbookPageDetailsModuleScss.fullRow}>
                  <div className={EhsHandbookPageDetailsModuleScss.imgCol}><img src={this.props.logoUrl} alt='Logo' /></div>
                  <div className={EhsHandbookPageDetailsModuleScss.detailCol}>
                    <ul className={EhsHandbookPageDetailsModuleScss.listItem}>
                      {selectedPageProperties.map((item) => {
                        return (<li>
                          <span>{item.key} : </span> {item.text}
                        </li>);
                      })}
                    </ul>
                  </div>
                </div>
              }
              {(this.props.layoutType === 'basic') &&
                <div>
                  {selectedPageProperties.map((item) => {
                    return (<div>
                      {(item.type === 'TaxonomyFieldType') &&
                        <div className={EhsHandbookPageDetailsModuleScss.column}>
                          <span className={EhsHandbookPageDetailsModuleScss.title}>{item.key}:</span>
                          <span className={EhsHandbookPageDetailsModuleScss.subTitle}>{item.text}</span>
                        </div>
                      }
                      {(item.type === 'TaxonomyFieldTypeMulti') &&
                        <div className={EhsHandbookPageDetailsModuleScss.column}>
                          <span className={EhsHandbookPageDetailsModuleScss.title}>{item.key}:</span>
                          {item.text ? item.text.split(', ').map((tag) => { return (<span className={EhsHandbookPageDetailsModuleScss.subTitle}>{tag}</span>); }) : item.text}
                        </div>
                      }
                    </div>);
                  })}

                  {this.props.pageProperties.indexOf(literalVersion) !== -1 &&
                    <div className={EhsHandbookPageDetailsModuleScss.column}>
                      <span className={EhsHandbookPageDetailsModuleScss.title}>Version:</span>
                      <span className={EhsHandbookPageDetailsModuleScss.subTitle}>{fieldValue[literalVersion]}</span>
                    </div>
                  }

                  {this.props.pageProperties.indexOf(literalScope) !== -1 &&
                    <div className={EhsHandbookPageDetailsModuleScss.column}>
                      <span className={EhsHandbookPageDetailsModuleScss.title}>Scope:</span>
                      <span className={EhsHandbookPageDetailsModuleScss.subTitle}>{fieldValue[literalScope]}</span>
                    </div>
                  }

                  {this.props.pageProperties.indexOf(literalModified) !== -1 &&
                    <div className={EhsHandbookPageDetailsModuleScss.column}>
                      <span className={EhsHandbookPageDetailsModuleScss.title}>Modified:</span>
                      <span className={EhsHandbookPageDetailsModuleScss.subTitle}>{fieldValue[literalModified] !== undefined ? fieldValue[literalModified].substr(0, fieldValue[literalModified].indexOf('T')) : '---'}</span>
                    </div>
                  }

                  {this.props.pageProperties.indexOf(literalCreated) !== -1 &&
                    <div className={EhsHandbookPageDetailsModuleScss.column}>
                      <span className={EhsHandbookPageDetailsModuleScss.title}>Created:</span>
                      <span className={EhsHandbookPageDetailsModuleScss.subTitle}>{fieldValue[literalModified] !== undefined ? fieldValue[literalCreated].substr(0, fieldValue[literalCreated].indexOf('T')) : '---'}</span>
                    </div>
                  }

                  {selectedPageProperties.map((item) => {
                    return (<div>
                      {(item.type !== 'TaxonomyFieldType' && item.type !== 'TaxonomyFieldTypeMulti' && item.key !== 'Version' && item.key !== 'Scope' && item.key !== 'Modified' && item.key !== 'Created') &&
                        <div className={EhsHandbookPageDetailsModuleScss.column}>
                          <span className={EhsHandbookPageDetailsModuleScss.title}>{item.key}:</span>
                          <span className={EhsHandbookPageDetailsModuleScss.subTitle}>{item.text}</span>
                        </div>
                      }
                    </div>);
                  })}
                </div>
              }
            </div>
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