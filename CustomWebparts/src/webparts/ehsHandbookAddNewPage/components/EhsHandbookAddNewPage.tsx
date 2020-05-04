import * as React from 'react';
import EhsHandbookAddNewPageModuleScss from './EhsHandbookAddNewPage.module.scss';
import { IEhsHandbookAddNewPageProps } from './IEhsHandbookAddNewPageProps';
import { IEhsHandbookAddNewPageState } from './IEhsHandbookAddNewPageState';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Panel } from 'office-ui-fabric-react';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';
import { HandbookComposite } from '../../../BAL/HandbookComposite';

export default class EhsHandbookAddNewPage extends React.Component<IEhsHandbookAddNewPageProps, IEhsHandbookAddNewPageState> {
  private thumbnailImage: File;
  private bannerImage: File;
  private options: IDropdownOption[] = [];

  constructor(props: IEhsHandbookAddNewPageProps) {
    super(props);

    this.state = {
      showPanel: false,
      pageTemplateUrl: '',
      pageName: '',
      topicPageLayout: '',
      pageNameErrorMessage: '',
      subjectDescription: '',
      subjectImage: '',
      subjectImageErrorText: '',
      subjectBannerImage: '',
      subjectBannerImageErrorText: '',
      loading: false,
      scope: (this.props.pageScopes.length > 0) ? (this.props.parentPageScope !== null ? this.props.pageScopes.filter(pageScope => pageScope.key === this.props.parentPageScope)[0] : this.props.pageScopes[0]) : { key: 'External', text: 'External' }
    };

    if (this.props.pageScopes.length > 0) {
      switch (this.props.parentPageScope) {
        case this.props.pageScopes[1] && this.props.pageScopes[1].key:
          this.options = this.props.pageScopes;
          this.options[0].disabled = true;
          break;
        case this.props.pageScopes[2] && this.props.pageScopes[2].key:
          this.options = this.props.pageScopes;
          this.options[0].disabled = true;
          this.options[1].disabled = true;
          break;
        default:
          this.options = this.props.pageScopes;
          break;
      }
    }
    this.validateFields = this.validateFields.bind(this);
    this.onConfigure = this.onConfigure.bind(this);

    this.uploadImages = this.uploadImages.bind(this);
    this.validateImageExtension = this.validateImageExtension.bind(this);
  }

  public componentDidUpdate(prevProps: IEhsHandbookAddNewPageProps) {
    if ((JSON.stringify(prevProps.pageScopes) !== JSON.stringify(this.props.pageScopes)) || (prevProps.parentPageScope !== this.props.parentPageScope)) {
      this.setState({ scope: (this.props.pageScopes.length > 0) ? (this.props.parentPageScope !== null ? this.props.pageScopes.filter(pageScope => pageScope.key === this.props.parentPageScope)[0] : this.props.pageScopes[0]) : { key: 'External', text: 'External' } });
      switch (this.props.parentPageScope) {
        case this.props.pageScopes[1] && this.props.pageScopes[1].key:
          this.options = this.props.pageScopes;
          this.options[0].disabled = true;
          break;
        case this.props.pageScopes[2] && this.props.pageScopes[2].key:
          this.options = this.props.pageScopes;
          this.options[0].disabled = true;
          this.options[1].disabled = true;
          break;
        default:
          this.options = this.props.pageScopes;
          break;
      }
    }
  }

  /**
   * Validation check for adding new Subject, Chapter or Topic
   */
  private validateFields(): void {
    let compositeContext = new HandbookComposite(this.props.context);
    if (this.state.pageName === '') {
      this.setState({ pageNameErrorMessage: 'Please enter the ' + this.props.pageType + ' Name' });
    } else {
      if (this.props.pageType === 'Subject') {
        let thumbnailImageOk = this.validateImageExtension(this.thumbnailImage, 'thumbnailImage');
        let bannerImageOk = this.validateImageExtension(this.bannerImage, 'bannerImage');
        if (thumbnailImageOk && bannerImageOk) {
          this.uploadImages(this.thumbnailImage, this.bannerImage).then(() => {
            try {
              this.setState({ loading: true });
              compositeContext.createNewPage(this.state.pageName, this.state.topicPageLayout, this.props.context.pageContext.listItem.id, this.props, this.state, this.state.scope.key, this.props.selectedList);
            } catch (error) {
              console.log(error);
            }
          });
        }
      } else {
        this.setState({ loading: true });
        compositeContext.createNewPage(this.state.pageName, this.state.topicPageLayout, this.props.context.pageContext.listItem.id, this.props, this.state, this.state.scope.key, this.props.selectedList);
      }
    }
  }

  /**
   * Footer content for Add a Page Panel
   */
  private onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={() => this.validateFields()} style={{ marginRight: '8px' }}>
          Save
        </PrimaryButton>
        <DefaultButton onClick={() => { this.setState({ showPanel: false, pageName: '', pageNameErrorMessage: '', subjectDescription: '', subjectImage: '', subjectBannerImage: '', subjectImageErrorText: '', subjectBannerImageErrorText: '' }); }}>
          Cancel
        </DefaultButton>
      </div>
    );
  }

  /**
   * Opens the web part property pane when Configure button of Placeholder control is clickec in page's Edit Mode
   */
  private onConfigure(): void {
    this.props.context.propertyPane.open();
  }

  /**
   * Check if image is of the extension .jpg, .png or .jpeg only
   *
   * @param imageFile Selected Image
   * @param imageFor Selected Image for thumbnail or banner
   */
  private validateImageExtension(imageFile: File, imageFor: string): boolean {
    if (imageFile === null || imageFile === undefined) {
      return true;
    } else {
      if ((imageFile.name.substr(imageFile.name.length - 4).toLowerCase() === '.png') || (imageFile.name.substr(imageFile.name.length - 4).toLowerCase() === '.jpg') || (imageFile.name.substr(imageFile.name.length - 5).toLowerCase() === '.jpeg')) {
        imageFor === 'thumbnailImage' ? this.thumbnailImage = imageFile : this.bannerImage = imageFile;
        return true;
      } else {
        imageFor === 'thumbnailImage' ? this.setState({ subjectImageErrorText: 'Please upload an image file only' }) : this.setState({ subjectBannerImageErrorText: 'Please upload an image file only' });
        return false;
      }
    }
  }

  /**
   * Upload the selected thumbnail and banner images in SiteAssets/HandbookImages folder
   *
   * @param thumbnailImage Selected Thumbnail Image
   * @param bannerImage Selected Banner Image
   */
  private async uploadImages(thumbnailImage: File, bannerImage: File) {
    let subjectImage: any;
    let subjectBannerImage: any;
    let subjectImageUrl = '';
    let subjectBannerImageUrl = '';
    let handbookComposite = new HandbookComposite(this.props.context);
    try {
      if (thumbnailImage !== null && thumbnailImage !== undefined) {
        subjectImage = await handbookComposite.uploadFiles(this.props.context.pageContext.web.serverRelativeUrl, thumbnailImage);
        subjectImageUrl = subjectImage.data.ServerRelativeUrl;
      }
      if (bannerImage !== null && bannerImage !== undefined) {
        subjectBannerImage = await handbookComposite.uploadFiles(this.props.context.pageContext.web.serverRelativeUrl, bannerImage);
        subjectBannerImageUrl = subjectBannerImage.data.ServerRelativeUrl;
      }
    } catch (error) {
      console.log('HandbookAddNewPage.uploadImages' + error);
    }

    this.setState({ subjectBannerImage: subjectBannerImageUrl, subjectImage: subjectImageUrl });
  }

  public render(): React.ReactElement<IEhsHandbookAddNewPageProps> {
    if (this.props.configured) {
      return (
        <div className={EhsHandbookAddNewPageModuleScss.ehsHandbookAddNewPage}>
          <div className={EhsHandbookAddNewPageModuleScss.container}>
            <div className={EhsHandbookAddNewPageModuleScss.row}>
              <div className={EhsHandbookAddNewPageModuleScss.column}>
                <DefaultButton text={'Add ' + this.props.pageType}
                  onClick={() => { this.setState({ showPanel: true }); }}
                  styles={{ root: { padding: '10px' } }}
                  iconProps={{ iconName: 'CalculatorAddition', styles: { root: EhsHandbookAddNewPageModuleScss.calculatorAdditionIcon } }}
                />
                <Panel headerText={'Add ' + this.props.pageType}
                  isOpen={this.state.showPanel}
                  onDismiss={() => { this.setState({ showPanel: false, pageName: '', pageNameErrorMessage: '', subjectDescription: '', subjectImage: '', subjectBannerImage: '', subjectImageErrorText: '', subjectBannerImageErrorText: '' }); }}
                  onRenderFooterContent={this.onRenderFooterContent}
                >
                  <Dropdown label='Scope'
                    defaultSelectedKey={this.state.scope.key}
                    options={this.options}
                    onChanged={(val) => { this.setState({ scope: val }); }}
                  />
                  <TextField required={true} errorMessage={this.state.pageNameErrorMessage} label={this.props.pageType + ' Name'} value={this.state.pageName} onChanged={(value) => { this.setState({ pageName: value }); }} />
                  {this.props.pageType === 'Subject' &&
                    <TextField label='Description'
                      value={this.state.subjectDescription}
                      onChanged={(value) => { if (value.length <= 120) { this.setState({ subjectDescription: value }); } else { this.setState({ subjectDescription: this.state.subjectDescription }); } }}
                      multiline={true}
                      placeholder='Please enter upto 120 characters only'
                    />
                  }
                  {this.props.pageType === 'Subject' &&
                    <div style={{ marginTop: '5px' }}>
                      <label title='Thumbnail Image' >Thumbnail Image</label>
                      <input style={{ marginTop: '5px' }} type='file' onChange={(value) => { this.thumbnailImage = value.target.files.item(0); }} accept='image/*' />
                      <p style={{ color: '#a80000', fontSize: '12px', marginTop: '5px', marginBottom: '5px' }}>{this.state.subjectImageErrorText}</p>
                    </div>
                  }
                  {this.props.pageType === 'Subject' &&
                    <div style={{ marginTop: '5px' }}>
                      <label title='Banner Image' >Banner Image</label>
                      <input style={{ marginTop: '5px' }} type='file' onChange={(value) => { this.bannerImage = value.target.files.item(0); }} accept='image/*' />
                      <p style={{ color: '#a80000', fontSize: '12px', marginTop: '5px', marginBottom: '5px' }}>{this.state.subjectBannerImageErrorText}</p>
                    </div>
                  }
                  <br />
                  {this.state.loading === true &&
                    <Spinner label={'Creating New ' + this.props.pageType} ariaLive='assertive' />
                  }
                </Panel>
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