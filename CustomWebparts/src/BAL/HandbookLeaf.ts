import { HandbookContentType } from './HandbookContentType';
import { IHandbookPageType } from './IHandbookPageType';
import { SPDataOperations } from '../DAL/SPDataOperations';
import { WebPartContext } from '@microsoft/sp-webpart-base';
export class HandbookLeaf implements IHandbookPageType {
    public id: number;
    public title: string;
    public link: string;
    public parentId: number;
    public scope: string;
    public bannerImageUrl: string;
    public description: string;
    public contentType: HandbookContentType;
    public spContext: WebPartContext;

    constructor(context: WebPartContext) {
        this.spContext = context;
    }

    public async getPageDetails(selectedList: string, pageId: number): Promise<IHandbookPageType> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            let pageContext = await oSPDataOperations.getPageDetails(selectedList, pageId);

            let allFieldsFromContentType = new HandbookContentType(this.spContext);
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
            console.log('HandbookLeaf.getPageDetails: ' + error);
        }
    }

}