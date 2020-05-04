import * as React from 'react';
import { Callout, getId, IIconProps, ActionButton } from 'office-ui-fabric-react';
import * as strings from 'EhsHandbookListingWebPartStrings';
import EhsHandbookListingModuleScss from '../EhsHandbookListing.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICalloutProps {
    additionalColumn: string[];
    currentItem: any;
    context: WebPartContext;
}

export interface ICalloutState {
    isCalloutVisible?: boolean;
}

export default class IconCallout extends React.Component<ICalloutProps, ICalloutState> {
    public state: ICalloutState = {
        isCalloutVisible: false
    };

    private menuButtonElement: any = React.createRef<HTMLDivElement>();
    private labelId: string = getId('callout-label');
    private descriptionId: string = getId('callout-description');
    private emojiIcon: IIconProps = { iconName: strings.callOutIconType };
    public render(): JSX.Element {
        let columnDetails: any;
        let additionalColumnsCount: number = 0;
        if (!!this.props.additionalColumn && this.props.additionalColumn !== null) {
            if (this.props.additionalColumn.length > 0) {
                additionalColumnsCount = this.props.additionalColumn.length;
            }
        }
        return (
            <div>
                <div className={EhsHandbookListingModuleScss.buttonArea} ref={this.menuButtonElement}>
                    <ActionButton className={EhsHandbookListingModuleScss.callOutButtonStyle} onClick={this.onShowMenuClicked} iconProps={this.emojiIcon}>
                        View Details
                    </ActionButton>
                </div>
                {this.state.isCalloutVisible && (
                    <Callout
                        className={EhsHandbookListingModuleScss.callout}
                        ariaLabelledBy={this.labelId}
                        ariaDescribedBy={this.descriptionId}
                        role={strings.callOutRole}
                        gapSpace={0}
                        target={this.menuButtonElement.current}
                        onDismiss={this.onCalloutDismiss}
                        setInitialFocus={true}
                    >
                        {(additionalColumnsCount > 0) ?
                            <div className={EhsHandbookListingModuleScss.inner}>{
                                this.props.additionalColumn.map((colName) => {
                                    columnDetails = this.props.currentItem.contentType.columns.filter(col => col.name === colName)[0];
                                    if (columnDetails !== undefined) {
                                        return (<div><span className={EhsHandbookListingModuleScss.heading}>{colName}</span><span> : {this.getColumnValue(columnDetails)} </span></div>);
                                    }
                                })}</div> : <div></div>
                        }
                    </Callout>
                )}
            </div>
        );
    }

    private onShowMenuClicked = (): void => {
        this.setState({
            isCalloutVisible: !this.state.isCalloutVisible
        });
    }

    private onCalloutDismiss = (): void => {
        this.setState({
            isCalloutVisible: false
        });
    }

    private formatDate(dateToFormat: string): string {
        let formattedDate: string = null;
        let months: string[] = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        try {
            if (dateToFormat !== null) {
                let dateValue: Date = new Date(dateToFormat);
                let dateMonth = dateValue.getMonth();
                let dateDay = dateValue.getDate();
                let dateYear = dateValue.getFullYear();
                formattedDate = dateDay + '-' + months[dateMonth] + '-' + dateYear;
            }
            return formattedDate;
        } catch (error) {
            console.log(error);
            return formattedDate;
        }
    }

    private getColumnValue(columnDetails: any): string {
        let columnValue: string = strings.emptyString;
        try {
            switch (columnDetails.columnType) {
                case strings.colTypeDateTime:
                    columnValue = this.formatDate(columnDetails.value);
                    break;
                case strings.colTypeUserMulti:
                    let userNames: string[] = [];
                    if (columnDetails.value !== undefined && columnDetails.value !== null) {
                        columnDetails.value.map((userName) => {
                            userNames.push(userName.Title);
                        });
                        columnValue = userNames.join(`,`);
                    } else { columnValue = strings.emptyString; }
                    break;
                case strings.colTypeTaxonomyFieldType:
                    columnValue = (!!columnDetails.value) ? columnDetails.value.Label : strings.emptyString;
                    break;
                case strings.colTypeTaxonomyFieldTypeMulti:
                    let taxonomyValues: string[] = [];
                    columnDetails.value.map((taxonomyValue) => {
                        taxonomyValues.push(taxonomyValue.Label);
                    });
                    columnValue = taxonomyValues.join(`,`);
                    break;
                case strings.colTypeText:
                    if ((columnDetails.value).indexOf(strings.principalNameIntials) !== -1) {
                        columnValue = (columnDetails.value).split(strings.principalNameIntials)[1];
                    } else {
                        columnValue = columnDetails.value;
                    }
                    break;
                case strings.colTypeLookup:
                    columnValue = (!!columnDetails) ? (!!columnDetails.value ? columnDetails.value.Id : strings.emptyString) : strings.emptyString;
                    break;
                default:
                    columnValue = (!!columnDetails.value) ? (columnDetails.value).toString() : strings.emptyString;
            }
            return columnValue;
        } catch (error) {
            console.log(error);
        }
    }
}