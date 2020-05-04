import * as React from 'react';
import EhsHandbookNavigationModuleScss from './EhsHandbookNavigation.module.scss';
import { IEhsHandbookNavigationProps } from './IEhsHandbookNavigationProps';
import { Nav } from 'office-ui-fabric-react/lib/Nav';
import { IEhsHandbookNavigationState } from './IEhsHandbooknavigationState';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';
import { HandbookComposite } from '../../../BAL/HandbookComposite';

export default class EhsHandbookNavigation extends React.Component<IEhsHandbookNavigationProps, IEhsHandbookNavigationState> {
  public subjectPageId: number = 0;

  constructor(props: IEhsHandbookNavigationProps) {
    super(props);

    this.state = {
      allTopicsLinks: []
    };

    this.onConfigure = this.onConfigure.bind(this);
  }

  public async componentDidMount() {
    let compositePage = new HandbookComposite(this.props.context);
    this.subjectPageId = await compositePage.getRootNodeId(this.props.selectedList, this.props.context.pageContext.listItem.id);
    await compositePage.getChildNodes(this.props.selectedList, this.subjectPageId, null, this.props.context);
    this.setState({ allTopicsLinks: compositePage.allNavLink });
  }

  public async componentDidUpdate(prevProps: IEhsHandbookNavigationProps) {
    /* Render updated topics when the selected subject property value is updated in the web part*/
    if ((prevProps.selectedList !== this.props.selectedList)) {
      let compositePage = new HandbookComposite(this.props.context);
      this.subjectPageId = await compositePage.getRootNodeId(this.props.selectedList, this.props.context.pageContext.listItem.id);
      await compositePage.getChildNodes(this.props.selectedList, this.subjectPageId, null, this.props.context);
      this.setState({ allTopicsLinks: compositePage.allNavLink });
    }
  }

  /**
   * Opens the web part property pane when Configure button of Placeholder control is clickec in page's Edit Mode
   */
  private onConfigure(): void {
    this.props.context.propertyPane.open();
  }

  public render(): React.ReactElement<IEhsHandbookNavigationProps> {
    if (this.props.configured) {
      return (
        <div className={EhsHandbookNavigationModuleScss.ehsHandbookNavigation}>
          <div className={EhsHandbookNavigationModuleScss.container}>
            <div className={EhsHandbookNavigationModuleScss.row}>
              <div className={EhsHandbookNavigationModuleScss.column}>
                <Nav
                  groups={[
                    {
                      links: this.state.allTopicsLinks
                    }
                  ]}
                />
              </div>
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
