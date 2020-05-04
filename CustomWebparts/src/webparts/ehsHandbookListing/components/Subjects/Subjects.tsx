import * as React from 'react';
import EhsHandbookListingModuleScss from '../EhsHandbookListing.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { HandbookComposite } from '../../../../BAL/HandbookComposite';
import { Link } from 'office-ui-fabric-react';
import IconCallout from './IconCallout';
import * as strings from 'EhsHandbookListingWebPartStrings';

export interface ISubjectsProps {
  siblingMetadata: HandbookComposite[];
  context: WebPartContext;
  logsSiteUrl: string;
  logsTitle: string;
  writeToDebug: boolean;
  sortOrder: string;
  sortType: string;
  sortByType: string;
  sortByField: string;
  additionalColumn: string[];
}
export class Subjects extends React.Component<ISubjectsProps, any> {

  constructor(props: ISubjectsProps) {
    super(props);

    this.state = {
      allItems: []
    };

    this.createElement = this.createElement.bind(this);
  }

  public componentDidMount() {
    this.createElement();
  }

  public componentDidUpdate(prevProps: ISubjectsProps) {
    /* Render updated sibling when the siblingMetadata property value is updated from EhsHandbookListing.tsx component*/
    if ((prevProps.siblingMetadata !== this.props.siblingMetadata) || (prevProps.sortOrder !== this.props.sortOrder) || (prevProps.sortType !== this.props.sortType) || (prevProps.sortByType !== this.props.sortByType) || (prevProps.sortByField !== this.props.sortByField)) {
      this.createElement();
    }
  }

  /**
   * Dynamically create element for rendering all the Siblings
   */
  private createElement() {
    let additionalColumnsCount: number = 0;
      if (!!this.props.additionalColumn && this.props.additionalColumn !== null) {
        if (this.props.additionalColumn.length > 0) {
          additionalColumnsCount = this.props.additionalColumn.length;
        }
      }

    try {
      let allSiblings = [];
      let allSiblingElements = [];
      let columnCount: number = 0;
      let sortOrder: any = this.props.sortByField === '' ? this.props.sortOrder : this.props.sortByField;
      let sortType: any = this.props.sortByType === '' ? this.props.sortType : this.props.sortByType;

      let sortedSiblingMetadata: any = sortType === strings.ascendingOptionKey ?
        this.props.siblingMetadata.sort((item1, item2) => {
          let value1 = item1.contentType.columns[item1.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;
          let value2 = item2.contentType.columns[item2.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;
          return value1.localeCompare(value2, 'en-u-kn-true');
        })
        : this.props.siblingMetadata.sort((item1, item2) => {
          let value1 = item1.contentType.columns[item1.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;
          let value2 = item2.contentType.columns[item2.contentType.columns.map(e => e.internalName.toUpperCase()).indexOf(sortOrder.toUpperCase())].value;

          return value2.localeCompare(value1, 'en-u-kn-true');
        });
        sortedSiblingMetadata.map((item, key) => {
        let bannerImageUrl = item.bannerImageUrl === undefined ? '/_layouts/15/images/sitepagethumbnail.png' : item.bannerImageUrl;
        let technicalContact: any;
        if (item.contentType.columns.filter(column => column.name === 'Technical Contact').length !== 0) {
          technicalContact = item.contentType.columns.filter(column => column.name === 'Technical Contact')[0].value;
        }
        columnCount += 1;
        allSiblings.push(<div className={EhsHandbookListingModuleScss.column}>
          <div><a href={item.link} style={{ textDecoration: 'none', color: 'black' }}>
            <img className={EhsHandbookListingModuleScss.image} src={bannerImageUrl} alt='Banner Iamge' />
          </a></div>
          <div className={EhsHandbookListingModuleScss.siblingDetails}>
            <div className={EhsHandbookListingModuleScss.siblingTitle}>{item.title}</div>
            <div className={EhsHandbookListingModuleScss.technicalContact}>
              {
                technicalContact !== undefined &&
                <div>Technical Contact: {technicalContact.map((contactPerson) => {
                  return (<span><Link href={'mailto:' + contactPerson.EMail}>{contactPerson.Title}</Link></span>);
                })}</div>
              }
            </div>
            {additionalColumnsCount > 0 ?
              <div>
                <IconCallout additionalColumn={this.props.additionalColumn} currentItem={item} context={this.props.context}></IconCallout>
              </div>
              : <></>
            }
          </div>
        </div>);

        if (columnCount % 3 === 0) {
          allSiblingElements.push(<div className={EhsHandbookListingModuleScss.row}>{allSiblings[columnCount - 3]}{allSiblings[columnCount - 2]}{allSiblings[columnCount - 1]}</div>);
        }
      });
      if (columnCount % 3 === 1) {
        allSiblingElements.push(<div className={EhsHandbookListingModuleScss.row}>{allSiblings[columnCount - 1]}</div>);
      } else if (columnCount % 3 === 2) {
        allSiblingElements.push(<div className={EhsHandbookListingModuleScss.row}>{allSiblings[columnCount - 2]}{allSiblings[columnCount - 1]}</div>);
      }
      this.setState({ allItems: allSiblingElements });

    } catch (error) {
      console.log({ errorMessage: error.message, errorMethod: 'Siblings.createElement' });
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
