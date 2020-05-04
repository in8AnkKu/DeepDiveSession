import { SPDataOperations } from '../DAL/SPDataOperations';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { HandbookDefaultColumns } from './HandbookDefaultColumns';
export class HandbookColumn {
    public name: string; //During Get
    public columnType?: string;  // Will be called on requiremnet of type
    public internalName?: string;
    public value?: any; // During Get
    public spContext?: WebPartContext;
    constructor(context?: WebPartContext) {
        this.spContext = context;
    }

    public async getScopeChoices?(listId: string): Promise<any[]> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            return await oSPDataOperations.loadScopeChoices(listId);
        } catch (error) {
            console.log('HandbookColumn.loadScopeChoices' + error);
        }
    }

    public async loadSingleTaxonomyValue?(lists: string, fields: string[], ids: string): Promise<any> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            return await oSPDataOperations.loadSingleTaxonomyValue(lists, fields, ids);
        } catch (error) {
            console.log('HandbookColumn.loadSingleTaxonomyValue: ' + error);
        }
    }

    public async getAllColumnsForContentType?(lists: string, pageId: number): Promise<any> {
        try {
            let oSPDataOperations = new SPDataOperations(this.spContext);
            return await oSPDataOperations.getAllColumnsForContentType(lists, pageId);
        } catch (error) {
            console.log('HandbookColumn.getAllColumnsForContentType: ' + error);
        }
    }
}