import { sp, PermissionKind } from '@pnp/sp';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-webpart-base';
//Polyfill to fix IE issues
import '@pnp/polyfill-ie11';
import { ClientSidePage, ClientSideWebpart } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

export class SPDataOperations {
  public context: WebPartContext;
  constructor(spContext: WebPartContext) {
    this.context = spContext;

    sp.setup({
      spfxContext: this.context
    });
  }

  /**
   * Load all Subjects in the Handbook for the Select Subject dropdown
   *
   * @param listId The list selected in web part property
   */
  public async loadRootLevelData(lists: string, contentTypeID: string): Promise<any> {
    let allSubjects: any;
    try {
      allSubjects = await sp.web.lists.getById(lists).items.select('PageId', 'Title').filter(`ContentTypeId eq '` + contentTypeID + `'`).get();
    } catch (error) {
      console.log('SPDataOperations.loadRootLevelData: ' + error);
    }
    return allSubjects;
  }

  public async loadSingleTaxonomyValue(lists: string, fields: string[], ids: string): Promise<any> {
    let termLabel: any[];
    try {
      termLabel = await sp.web.lists.getByTitle(lists).items.select(...fields).filter(ids).get();
    } catch (error) {
      console.log('SPDataOperations.loadSingleTaxonomyValue: ' + error);
    }
    return termLabel;
  }
  /**
   * Get all the content types for the selected list
   *
   * @param lists The list selected in web part property for which content types are needed
   */
  public async loadContentTypes(lists: string): Promise<IPropertyPaneDropdownOption[]> {
    let allContentTypeNames: IPropertyPaneDropdownOption[] = [];
    try {
      let allContentTypesItems: any = await sp.web.lists.getById(lists).contentTypes.select('Name').get();
      allContentTypesItems.map((currentContentType: any) => {
        allContentTypeNames.push({ key: currentContentType.Name, text: currentContentType.Name });
      });
    } catch (error) {
      console.log('SPDataOperations.loadContentTypes: ' + error);
    }
    return allContentTypeNames;
  }
  /**
   * Gets the required field values of Managed Metadata type list columns
   *
   * @param lists The list for which the fields of type Managed Metadata need to be retrieved
   * @param fields The Managed Metadata fields which need to be retrieved
   */
  public async loadTermSetFields(lists: string, fields: string[]): Promise<any> {
    let termSetFields: any[];
    try {
      termSetFields = await sp.web.lists.getById(lists).fields.select(...fields).filter(`TypeDisplayName eq 'Managed Metadata'`).get();
    } catch (error) {
      console.log('SPDataOperations.loadTermSetFields: ' + error);
    }
    return termSetFields;
  }
  /**
   * Gets the available Choices in the Scope Choice field
   *
   * @param lists The list for which the fields of type Managed Metadata need to be retrieved
   */
  public async loadScopeChoices(lists: string): Promise<any[]> {
    try {
      return await sp.web.lists.getById(lists).fields.select('Choices').filter(`InternalName eq 'Scope'`).get();
    } catch (error) {
      console.log('SPDataOperations.loadScopeChoices: ' + error);
    }
  }
  /**
  * Check if the current user has requested permissions on a list
  * @param lists The list on which user permission needs to be checked
  * @param permissionKey The permission kind for which user needs to be authorized
  */
  public async checkUserPermissions(lists: string, permissionKey: string): Promise<boolean> {
    let perms: boolean = false;
    let permission: PermissionKind = PermissionKind[permissionKey];
    try {
      let effectiveBasePermissions = await sp.web.lists.getById(lists).effectiveBasePermissions.get();
      perms = await sp.web.lists.getById(lists).hasPermissions(effectiveBasePermissions, permission);
    } catch (error) {
      console.log('SPDataOperations.checkUserPermissions: ' + error);
    }
    return perms;
  }

  public async loadPageContentType(lists: string, pageId: number, internalName: any, expandFields: any): Promise<any> {
    let contentType: any[];
    try {
      contentType = await sp.web.lists.getById(lists).items.filter('ID eq ' + pageId).select(...internalName).expand(...expandFields).get();
    } catch (error) {
      console.log('SPDataOperations.loadPageContentType: ' + error);
    }
    return contentType;
  }

  public async getAllColumnsForContentType(lists: string, pageId: number): Promise<any> {
    let fields: string[] = ['Title', 'TypeAsString', 'InternalName', 'DependentLookupInternalNames'];
    let contentTypeField: any[];
    let contentTypeId: string;
    try {
      let getContentTypeId = await sp.web.lists.getById(lists).items.select('ContentTypeId').filter('ID eq ' + pageId).get();
      if (getContentTypeId.length > 0) {
        getContentTypeId.map((item) => {
          contentTypeId = item.ContentTypeId;
        });
        contentTypeField = await sp.web.lists.getById(lists).contentTypes.getById(contentTypeId).fields.select(...fields).filter(`TypeDisplayName ne 'Computed' and TypeAsString ne 'LookupMulti' and TypeAsString ne 'HTML' and TypeAsString ne 'Guid' and TypeAsString ne 'Note' and TypeAsString ne 'URL' and (Group ne '_Hidden' or InternalName eq 'Title')`).orderBy('Title').get();
      }
    } catch (error) {
      console.log('SPDataOperations.getAllColumnsForContentType: ' + error);
    }
    return contentTypeField;
  }

  public async getChildNodes(selectedList: string, parentId: number): Promise<any[]> {
    try {
      return new Promise<any[]>(async (resolve) => {
        let childNodes: any[];
        childNodes = await sp.web.lists.getById(selectedList).items.filter(`ParentId eq ` + parentId).select('*', 'EncodedAbsUrl').get();
        resolve(childNodes);
      });
    } catch (error) {
      console.log('SPDataOperations.getChildNodes: ' + error);
    }
  }

  public static DELETECHILDNODE() {
    return;
  }

  public async getContentType(listName: any, pageId: any): Promise<any> {
    try {
      let contentTypeId: any;
      let getContentTypeId = await sp.web.lists.getById(listName).items.select('ContentTypeId').filter('ID eq ' + pageId).get();
      getContentTypeId.map((item) => {
        contentTypeId = item.ContentTypeId;
      });

      let contentTypeContextfromId = await sp.web.lists.getById(listName).contentTypes.filter(`StringId eq '` + contentTypeId + `'`).get();
      return Promise.resolve(contentTypeContextfromId);
    } catch (error) {
      console.log('SPDataOperations.getContentType: ' + error);
    }
  }

  /**
   * Loads current page's metadata to be rendered
   * Retrieved data stored in global variable for reuse across multiple instances of the web part on the same page
   */
  public async getPageDetails(listId: string, pageID: number): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let currentPageContext = await sp.web.lists.getById(listId).items.getById(pageID).select('Title', '*', 'OData__UIVersionString', 'Scope', 'Modified', 'Created', 'EncodedAbsUrl', 'FileRef', 'Id').get();
        resolve(currentPageContext);
      });
    } catch (error) {
      console.log('SPDataOperations.getPageDetails: ' + error);
    }
  }

  public async getPageDetailsByUrl(selectedList: any, newPageName: any): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let currentPageContext = await sp.web.lists.getById(selectedList).items.select('FileRef', 'Id').filter(`FileRef eq '` + '/sites' + '/SitePages/' + newPageName.split(' ').join('-') + `.aspx'`).get();
        resolve(currentPageContext);
      });
    } catch (error) {
      console.log('SPDataOperations.getPageDetailsByUrl: ' + error);
    }
  }

  public async getHandbookContentTypes(listId: string): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let handbookContentTypes = await sp.web.lists.getById(listId).select('group', 'Name', 'Id').contentTypes.filter(`substringof('Handbook',group)`).get();
        resolve(handbookContentTypes);
      });
    } catch (error) {
      console.log('SPDataOperations.getHandbookContentTypes: ' + error);
    }
  }

  public async getAllItemsWithNoParent(listId: string): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let rootLevelItems = await sp.web.lists.getById(listId).items.select('Parent/ID', 'ContentTypeId', 'Id').expand('Parent').filter(`Parent/ID eq null`).get();
        resolve(rootLevelItems);
      });
    } catch (error) {
      console.log('SPDataOperations.getAllItemsWithNoParent: ' + error);
    }
  }

  public async getParentContentType(listId: string, childContentTypeId: string): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let parentContentType = await sp.web.lists.getById(listId).contentTypes.getById(childContentTypeId).parent.select('Id').get();
        resolve(parentContentType);
      });
    } catch (error) {
      console.log('SPDataOperations.getParentContentType: ' + error);
    }
  }

  public async getFileContext(oldPageUrl: string): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let fileContext = await ClientSidePage.fromFile(sp.web.getFileByServerRelativeUrl(oldPageUrl));
        resolve(fileContext);
      });
    } catch (error) {
      console.log('SPDataOperations.getFileContext: ' + error);
    }
  }

  public async getExistingPageContext(serverRelativeUrl: string, finalPageName: string): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let existingPageContext = await sp.web.getFolderByServerRelativeUrl(`${serverRelativeUrl}/SitePages/`).files.filter(`Name eq '${finalPageName.replace(/ /g, '-')}.aspx'`).get();
        resolve(existingPageContext);
      });
    } catch (error) {
      console.log('SPDataOperations.getExistingPageContext: ' + error);
    }
  }

  public async getListsContext(selectedList: string): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let listContext = await sp.web.lists.getById(selectedList);
        resolve(listContext);
      });
    } catch (error) {
      console.log('SPDataOperations.getListsContext: ' + error);
    }
  }

  public async addClientSidePage(finalPageName: string): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let clientSidePage = await sp.web.addClientSidePage(finalPageName + '.aspx');
        resolve(clientSidePage);
      });
    } catch (error) {
      console.log('SPDataOperations.addClientSidePage: ' + error);
    }
  }

  public async updatePageMetaData(selectedList: any, newContentTypeId: string, parentPageId: any, scope: any, newPageItemId: any): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let clientSidePage = await sp.web.lists.getById(selectedList).items.getById(newPageItemId).update({
          ContentTypeId: newContentTypeId,
          ParentId: parentPageId,
          Scope: scope,
          PageId: newPageItemId
        });
        resolve(clientSidePage);
      });
    } catch (error) {
      console.log('SPDataOperations.updatePageMetaData: ' + error);
    }
  }

  public async movePage(newPageItemRelativeUrl: string, scope: any) {
    try {
      return await sp.web.getFileByServerRelativeUrl(newPageItemRelativeUrl).moveTo(newPageItemRelativeUrl.replace(`/SitePages/`, `/SitePages/${scope}/`));
    } catch (error) {
      console.log('SPDataOperations.movePage: ' + error);
    }
  }

  public async getUserProfile(): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let userProfile = await sp.profiles.myProperties.get();
        resolve(userProfile);
      });
    } catch (error) {
      console.log('SPDataOperations.getUserProfile: ' + error);
    }
  }

  public async getPageData(listId: string, fields: string[], filter: string): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let pageData = sp.web.lists.getById(listId).items.select(...fields).filter(filter + ' eq 1').get();
        resolve(pageData);
      });
    } catch (error) {
      console.log('SPDataOperations.getPageData: ' + error);
    }
  }

  public async getListItem(listId: string, itemId: number): Promise<any> {
    try {
      return await sp.web.lists.getById(listId).items.getById(itemId).get();
    } catch (error) {
      console.log('SPDataOperations.getListItem: ' + error);
    }
  }

  public async uploadFiles(siteUrl: string, image: File): Promise<any> {
    try {
      return await sp.web.getFolderByServerRelativeUrl(`${siteUrl}/SiteAssets/HandbookImages`).files.add(image.name, image);
    } catch (error) {
      console.log('SPDataOperations.uploadFiles: ' + error);
    }
  }

  public async checkOutPage(contextSPHTTPClient: any, absoluteUrl: string, newPageItemId: any): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let checkedOutPageOutput = contextSPHTTPClient.post(
          `${absoluteUrl}/_api/sitepages/pages(${newPageItemId})/checkoutpage`,
          SPHttpClient.configurations.v1,
          {});
        resolve(checkedOutPageOutput);
      });
    } catch (error) {
      console.log(error);
    }
  }

  public async setPageProperties(contextSPHTTPClient: any, absoluteUrl: string, body: string, newPageItemId: any): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let pagePropertiesOutput = contextSPHTTPClient.post(
          `${absoluteUrl}/_api/sitepages/pages(${newPageItemId})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE'
            },
            body: body
          });
        resolve(pagePropertiesOutput);
      });
    } catch (error) {
      console.log(error);
    }
  }

  public async publishPage(contextSPHTTPClient: any, absoluteUrl: string, newPageItemId: any): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let pagePublishedOutput = await contextSPHTTPClient.post(`${absoluteUrl}/_api/sitepages/pages(${newPageItemId})/publish`, SPHttpClient.configurations.v1, {});
        resolve(pagePublishedOutput);
      });
    } catch (error) {
      console.log(error);
    }
  }

  public async getLayoutWebpartsContent(contextSPHTTPClient: any, absoluteUrl: string, newPageItemId: any): Promise<any> {
    try {
      return new Promise(async (resolve) => {
        let pagePublishedOutput = await contextSPHTTPClient.get(`${absoluteUrl}/_api/sitepages/pages(${newPageItemId})?$select=LayoutWebpartsContent`, SPHttpClient.configurations.v1);
        resolve(pagePublishedOutput);
      });
    } catch (error) {
      console.log(error);
    }
  }

  public async addNewPage(templatePath: string, oldPage: any, selectedList: string, finalPageName: string, newPageName: string, parentId: number, scope: any, state: any, props: any, contentTypeLevel: number, contentTypeID: string) {
    let page: ClientSidePage;
    if (templatePath) {
      page = await oldPage.copyPage(await this.getListsContext(selectedList), finalPageName + '.aspx', newPageName, true);
    } else {
      page = await this.addClientSidePage(finalPageName);
    }
    let newPageItem: any;
    let newPage: any;
    try {
      newPage = page;
      newPage = await newPage.getItem();
      newPageItem = await this.getPageDetails(selectedList, newPage.Id);
    } catch (error) {
      newPageItem = await this.getPageDetailsByUrl(selectedList, newPageName);
    }
    let newPageItemId = newPageItem.Id;
    let newPageUrl = newPageItem.FileRef;
    let parentPageId = contentTypeLevel === 1 ? null : parentId;
    let newContentTypeId = contentTypeID;
    await this.updatePageMetaData(selectedList, newContentTypeId, parentPageId, scope, newPageItemId);
    let topLevelHeaderTopic = null;
    let newPageBannerImage = '';
    let topLevelImage: string = null;
    let topLevelDescription: string = null;
    let layoutWebpartsContentResponse: any;
    let layoutWebpartsContentValue;
    let layoutWebpartsContentJson;
    //We are using subject key word till the time we haven't made chnages in TSX for generic Handbook
    if (contentTypeLevel === 1) {
      newPageBannerImage = state.subjectBannerImage;
      topLevelImage = state.subjectImage !== '' ? state.subjectImage : `${props.context.pageContext.web.serverRelativeUrl}/_layouts/15/images/sitepagethumbnail.png`;
      topLevelDescription = state.subjectDescription;
    } else {
      layoutWebpartsContentResponse = await this.getLayoutWebpartsContent(props.context.spHttpClient, props.context.pageContext.web.absoluteUrl, parentPageId);
      layoutWebpartsContentValue = await layoutWebpartsContentResponse.json();
      layoutWebpartsContentJson = JSON.parse(layoutWebpartsContentValue.LayoutWebpartsContent);
      newPageBannerImage = layoutWebpartsContentJson[0].serverProcessedContent.imageSources.imageSource;
      topLevelHeaderTopic = layoutWebpartsContentJson[0].properties.title;
    }
    //Get current logged in user's profile properties and form a key value pair for setting up the Author People Web Part
    let myProps = await this.getUserProfile();
    let properties = {};
    myProps.UserProfileProperties.forEach((prop: any) => {
      properties[prop.Key] = prop.Value;
    });
    myProps.userProperties = properties;
    //Update user properties of Author People Web Part
    const authorWebPart = page.findControl<ClientSideWebpart>((csw: ClientSideWebpart) => {
      if (csw.title === 'People') {
        return (csw.data.webPartData.serverProcessedContent.searchablePlainTexts.title === 'Author');
      } else {
        return false;
      }
    });
    if (authorWebPart) {
      authorWebPart.data.webPartData.serverProcessedContent.searchablePlainTexts = {
        'title': 'Author',
        'persons[0].name': `${props.userProperties.PreferredName}`,
        'persons[0].email': `${props.userProperties.WorkEmail}`
      };
      authorWebPart.setProperties({
        persons: [
          {
            'id': `${props.userProperties.WorkEmail}`,
            'upn': '',
            'role': `${props.userProperties.Title}`,
            'department': `${props.userProperties.Department}`,
            'phone': '',
            'sip': ''
          }
        ]
      });
    }
    await page.save(false);
    this.updateLayoutWebpartsContent(props, newPageBannerImage, newPageName, topLevelHeaderTopic, newPageItemId, newPageUrl, topLevelImage, topLevelDescription, scope);
  }

  private async updateLayoutWebpartsContent(props: any, newPageBannerImage: string, newPageName: string, topLevelHeaderTopic: string, newPageItemId: number, newPageUrl: string, topLevelImage: string, topLevelDescription: string, scope: string | number) {
    try {
      //Default URL for the page banner image if none is selected
      if (newPageBannerImage === '') {
        newPageBannerImage = '/_LAYOUTS/IMAGES/SLEEKTEMPLATEIMAGETILE.JPG';
      }
      //LayoutWebPartsProperty value for new page
      let layoutJson = `[{
            'id': 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
            'instanceId': 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
            'title': 'Title area',
            'description': 'Title Region Description',
            'serverProcessedContent': {
              'htmlStrings': {},
              'searchablePlainTexts': {},
              'imageSources': {
                'imageSource': '${newPageBannerImage}'
              },
              'links': {},
              'customMetadata': {
                'imageSource': {
                  'siteId': 'null',
                  'webId': 'null',
                  'listId': 'null',
                  'uniqueId': 'null'
                }
              }
            },
            'dataVersion': '1.4',
            'properties': {
              'title': '${newPageName}',
              'imageSourceType': 2,
              'layoutType': 'FullWidthImage',
              'textAlignment': 'Left',
              'showTopicHeader': true,
              'showPublishDate': false,
              'topicHeader': '${topLevelHeaderTopic}',
              'authorByline': null,
              'authors': [],
              'altText': '',
              'translateX': 99.240698557327249,
              'translateY': 45.614035087719294,
              'webId': 'null',
              'siteId': 'null',
              'listId': 'null',
              'uniqueId': 'null'
            }
          }]`;

      const body: string = JSON.stringify({
        '__metadata': {
          'type': 'SP.Publishing.SitePage'
        },
        'TopicHeader': topLevelHeaderTopic,
        'BannerImageUrl': topLevelImage/* newPageBannerImage */,
        'Description': topLevelDescription,
        'LayoutWebpartsContent': layoutJson
      });

      /**
       * Checkout the new page and update banner properties using REST API
       */
      //let oSPDataOperations = new SPDataOperations(this.spContext);
      let data: any = await this.checkOutPage(props.context.spHttpClient, props.context.pageContext.web.absoluteUrl, newPageItemId);
      if (data.ok) {
        let pagePropertiesOutput = await this.setPageProperties(props.context.spHttpClient, props.context.pageContext.web.absoluteUrl, body, newPageItemId);

        if (pagePropertiesOutput.ok) {
          await this.publishPage(props.context.spHttpClient, props.context.pageContext.web.absoluteUrl, newPageItemId);
          let newPageItemRelativeUrl = newPageUrl.substring(newPageUrl.indexOf('/sites/'));
          await this.movePage(newPageItemRelativeUrl, scope);
          window.location.reload();
          let updatedPageUrl = newPageUrl.replace(`/SitePages/`, `/SitePages/${scope}/`);
          window.open(updatedPageUrl, '_blank');
        }
      }
    } catch (error) {
      console.log('HandbookPage.updateLayoutWebpartsContent: ' + error);
    }
  }
}
