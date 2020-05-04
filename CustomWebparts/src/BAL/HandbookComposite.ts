import { HandbookContentType } from './HandbookContentType';
import { IHandbookPageType } from './IHandbookPageType';
import { SPDataOperations } from '../DAL/SPDataOperations';
import { HandbookLeaf } from './HandbookLeaf';
import { INavLink } from './IHandbookNavLink';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { HandbookDefaultColumns } from './HandbookDefaultColumns';
export class HandbookComposite implements IHandbookPageType {
    public id: number;
    public title: string;
    public link: string;
    public parentId: number;
    public scope: string;
    public bannerImageUrl: string;
    public description: string;
    public contentType: HandbookContentType;
    public childNodes: IHandbookPageType[] = [];
    public allNavLink: INavLink[] = [];
    public itemId?: number;

    public spContext: WebPartContext;
    constructor(context: WebPartContext) {
        this.spContext = context;
    }

    public async getPageDetails(selectedList: string, pageId: number, spContext: WebPartContext): Promise<HandbookComposite> {
        try {
            if (this.spContext === undefined) {
                this.spContext = spContext;
            }
            let spContextObj = this.spContext;
            let oSPDataOperations = new SPDataOperations(spContextObj);
            let pageContext = await oSPDataOperations.getPageDetails(selectedList, pageId);

            let allFieldsFromContentType = new HandbookContentType(spContextObj);
            await allFieldsFromContentType.getAllColumnsForContentType(selectedList, pageId);
            this.title = pageContext.Title;
            this.link = pageContext.EncodedAbsUrl;
            this.id = pageContext.Id;
            this.parentId = pageContext.ParentId;
            this.scope = pageContext.Scope;
            this.contentType = allFieldsFromContentType;
            if (pageContext.BannerImageUrl) { this.bannerImageUrl = pageContext.BannerImageUrl.Url; }
            this.description = pageContext.Decsription;
            delete this.spContext;
            if (this.contentType) { delete this.contentType.spContext; }
            return this;
        } catch (error) {
            console.log('HandbookComposite.getPageDetails: ' + error);
        }
    }

    public async getHandBookNodes(listId: string, parentId: number, endLevel: number, spContext: WebPartContext) {
        try {
            if (this.spContext === undefined) { this.spContext = spContext; }
            let spContextObj = this.spContext;
            this.childNodes = [];
            let oHandBookcontentType = new HandbookContentType(this.spContext);
            let contentTypeDetails = await oHandBookcontentType.populateContentTypeLevel(listId);
            if (endLevel === null) {
                endLevel = contentTypeDetails.length;
            }
            if (parentId === null) {
                let oSPDataOperations = new SPDataOperations(this.spContext);
                let rootLevelItems = await oSPDataOperations.getAllItemsWithNoParent(listId);
                await Promise.all(rootLevelItems.map(async (rootLevelItem) => {

                    let handBookItem = contentTypeDetails.filter((handbookContentType) => handbookContentType.contentTypeID === rootLevelItem.ContentTypeId);
                    if (handBookItem.length > 0) {
                        await this.getParentDetails(listId, rootLevelItem.Id, spContextObj);
                        await this.getChildNodesDetails(rootLevelItem.Id, listId, endLevel, spContextObj);
                        delete this.spContext;
                        if (this.contentType) { delete this.contentType.spContext; }
                    }
                }));

            } else {
                await this.getParentDetails(listId, parentId, spContextObj);
                await this.getChildNodesDetails(parentId, listId, endLevel, spContextObj);
                delete this.spContext;
                if (this.contentType) { delete this.contentType.spContext; }
            }
        } catch (error) {
            console.log('HandbookComposite.getHandBookNodes: ' + error);
        }
    }

    private async getParentDetails(selectedList: string, itemId: number, spContext: WebPartContext) {
        try {
            if (this.spContext === undefined) { this.spContext = spContext; }
            let spContextObj = this.spContext;
            let oSPDataOperations = new SPDataOperations(this.spContext);
            let parentDetails = await oSPDataOperations.getPageDetails(selectedList, itemId);
            let childNodeDetails = await oSPDataOperations.getChildNodes(selectedList, itemId);
            if (childNodeDetails.length > 0) {
                let compositePage = new HandbookComposite(spContextObj);
                await compositePage.getPageDetails(selectedList, parentDetails.Id, spContextObj);
                this.childNodes.push(compositePage);
                let nodeNavComp: INavLink = await this.getNavLink(parentDetails);
                this.allNavLink.push(nodeNavComp);
            } else {
                let leafPage = new HandbookLeaf(spContextObj);
                await leafPage.getPageDetails(selectedList, parentDetails.Id);
                this.childNodes.push(leafPage);
                let nodeNavLeaf: INavLink = await this.getNavLink(parentDetails);
                this.allNavLink.push(nodeNavLeaf);
            }
        } catch (error) {
            console.log('HandbookComposite.getParentDetails: ' + error);
        }
    }

    private async getChildNodesDetails(parentId: number, selectedList: string, endLevel: number, spContext: WebPartContext) {
        try {
            if (this.spContext === undefined) { this.spContext = spContext; }
            let oSPDataOperations = new SPDataOperations(this.spContext);
            let spContextObj = this.spContext;
            let childNodeDetails = await oSPDataOperations.getChildNodes(selectedList, parentId);

            if (childNodeDetails.length > 0) {

                await Promise.all(childNodeDetails.map(async (singleNode) => {
                    let nextLevelChildNodes = await oSPDataOperations.getChildNodes(selectedList, singleNode.Id);
                    //There can be a case where it is a composite node but there is no child belonging to it
                    if (nextLevelChildNodes.length > 0) {
                        let compositePage = new HandbookComposite(spContextObj);
                        await compositePage.getPageDetails(selectedList, singleNode.Id, spContextObj);
                        if (compositePage.contentType.contentTypeLevel <= endLevel) {
                            this.childNodes.push(compositePage);
                            let nodeNavComp: INavLink = await this.getNavLink(singleNode);
                            this.allNavLink.push(nodeNavComp);
                            return this.getChildNodesDetails(singleNode.Id, selectedList, endLevel, spContextObj);
                        }
                    } else {
                        let leafPage = new HandbookLeaf(spContextObj);
                        await leafPage.getPageDetails(selectedList, singleNode.Id);
                        if (leafPage.contentType.contentTypeLevel <= endLevel) {
                            this.childNodes.push(leafPage);
                            let nodeNavLeaf: INavLink = await this.getNavLink(singleNode);
                            this.allNavLink.push(nodeNavLeaf);
                        }
                    }
                }));
            }
        } catch (error) {
            console.log('HandbookComposite.getChildNodesDetails: ' + error);
        }
    }

    private async getNavLink(node: any) {
        if (node.Scope === 'Internal') {
            return {
                name: node.Title,
                url: node.EncodedAbsUrl,
                parentId: node.ParentId,
                key: node.Id,
                isExpanded: true,
                icon: 'ProtectedDocument'
            };
        } else {
            return {
                name: node.Title,
                url: node.EncodedAbsUrl,
                parentId: node.ParentId,
                key: node.Id,
                isExpanded: true
            };
        }
    }

    public async getChildNodes(listId: string, parentId: number, endLevel: number, spContext: WebPartContext): Promise<HandbookComposite[]> {
        try {
            if (this.spContext === undefined) { this.spContext = spContext; }
            await this.getHandBookNodes(listId, parentId, endLevel, this.spContext);
            let arrayToTree = require('array-to-tree');
            let treeStructure = arrayToTree(this.childNodes, {
                parentProperty: 'parentId',
                childrenProperty: 'childNodes',
                customID: 'id'
            });
            let navLinkTreeStructure = arrayToTree(this.allNavLink, {
                parentProperty: 'parentId',
                childrenProperty: 'links',
                customID: 'key'
            });
            this.allNavLink = navLinkTreeStructure;
            return Promise.resolve(treeStructure);
        } catch (error) {
            console.log('HandbookComposite.getChildNodes: ' + error);
        }

    }

    public async getRootNodeId(list: any, itemId: any): Promise<number> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            let pageObject = await oSPDataOperations.getPageDetails(list, itemId);
            if ('ParentId' in pageObject) {
                if (pageObject.ParentId !== null) {
                    await this.getRootNodeId(list, pageObject.ParentId);
                } else {
                    this.itemId = pageObject.Id;
                }
                return this.itemId;
            }
        } catch (error) {
            console.log('HandbookComposite.getRootNodeId: ' + error);
        }
    }

    public async loadTermsetFields(listid: string, fields: string[]): Promise<any> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            return await oSPDataOperations.loadTermSetFields(listid, fields);
        } catch (error) {
            console.log('HandbookComposite.loadTermSetFields: ' + error);
        }
    }

    public async uploadFiles(siteUrl: string, image: File): Promise<any> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            return await oSPDataOperations.uploadFiles(siteUrl, image);
        } catch (error) {
            console.log('HandbookComposite.uploadFiles: ' + error);
        }
    }

    public async loadRootLevelData(listId: string, contentTypeID: string): Promise<any> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            return await oSPDataOperations.loadRootLevelData(listId, contentTypeID);
        } catch (error) {
            console.log('HandbookComposite.loadRootLevelData: ' + error);
        }
    }

    public async checkUserPermissions(lists: string, permissionKey: string): Promise<boolean> {
        let perms: boolean = false;
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            perms = await oSPDataOperations.checkUserPermissions(lists, permissionKey);
        } catch (error) {
            console.log('HandbookColumn.CheckUserPermissions: ' + error);
        }
        return perms;
    }

    public async loadContentTypes(lists: string): Promise<any> {
        let allContentTypeNames: any;
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            allContentTypeNames = await oSPDataOperations.loadContentTypes(lists);
        } catch (error) {
            console.log('HandbookComposite.loadContentTypes: ' + error);
        }
        return allContentTypeNames;
    }

    public async getListItem(listId: string, itemId: number): Promise<any> {
        let listItem: any;
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            listItem = await oSPDataOperations.getListItem(listId, itemId);
        } catch (error) {
            console.log('HandbookComposite.getListItem: ' + error);
        }
        return listItem;
    }

    public async createNewPage(newPageName: string, templatePath: string, parentId: number, props: any, state: any, scope: any, selectedList: string) {
        try {
            templatePath = props.templateUrl;
            let pageContentTypeLevel: number;
            let oSPDataOperations = new SPDataOperations(this.spContext);
            let parentContentTypeDetails = await oSPDataOperations.getContentType(selectedList, parentId);
            let contentTypeHandbook = new HandbookContentType(this.spContext);
            let contentTypeHandbookLevels = await contentTypeHandbook.populateContentTypeLevel(selectedList);
            if (parentContentTypeDetails[0].Group.indexOf('Handbook') < 0) {
                pageContentTypeLevel = 1;
            } else {
                let parentLevelInHandbookHierarchy = (await contentTypeHandbookLevels).filter((item) => (item.contentTypeID === parentContentTypeDetails[0].StringId))[0].contentTypeLevel;
                pageContentTypeLevel = (parentLevelInHandbookHierarchy + 1);
            }
            let contentType = (await contentTypeHandbookLevels).filter((item) => (item.contentTypeLevel === pageContentTypeLevel))[0];

            let oldPageUrl: string = templatePath;
            let sitesIndex: number = oldPageUrl.indexOf('/sites/');
            oldPageUrl = oldPageUrl.substring(sitesIndex);
            let oldPage: any;
            oldPage = await oSPDataOperations.getFileContext(oldPageUrl);

            let pageExists = true;
            let finalPageName = newPageName;
            let count: number = 1;
            while (pageExists === true) {
                let existingPage = await oSPDataOperations.getExistingPageContext(props.context.pageContext.web.serverRelativeUrl, finalPageName);
                if (existingPage[0]) {
                    pageExists = true;
                    finalPageName = newPageName + `${count}`;
                    count++;
                } else {
                    pageExists = false;
                }
            }
            await oSPDataOperations.addNewPage(templatePath, oldPage, selectedList, finalPageName, newPageName, parentId, scope, state, props, contentType.contentTypeLevel, contentType.contentTypeID);
        } catch (error) {
            console.log('CreateNewPage : ' + error);
        }
    }

    public async getPageData(listId: string): Promise<any> {
        try {
            let fields: string[] = [HandbookDefaultColumns.IS_TOPIC_TEMPLATE,
            HandbookDefaultColumns.ID,
            HandbookDefaultColumns.TITLE,
            HandbookDefaultColumns.ENCODEDABSURL,
            HandbookDefaultColumns.TOPIC_ICON,
            HandbookDefaultColumns.SELECTED_TOPIC_ICON];

            let oSPDataOperations = new SPDataOperations(this.spContext);
            return await oSPDataOperations.getPageData(listId, fields, HandbookDefaultColumns.IS_TOPIC_TEMPLATE);
        } catch (error) {
            console.log('HandbookPage.getPageData: ' + error);
        }
    }

    public async getPageDetailsById(listId: string, pageID: number): Promise<any> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            return await oSPDataOperations.getPageDetails(listId, pageID);
        } catch (error) {
            console.log(error);
        }
    }
}