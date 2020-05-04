import { HandbookColumn } from './HandbookColumn';
import { SPDataOperations } from '../DAL/SPDataOperations';
import { IHandbookContentType } from './IHandbookContentType';
import { WebPartContext } from '@microsoft/sp-webpart-base';
export class HandbookContentType {
  public name: string;
  public description: string;
  public columns: HandbookColumn[];
  public contentTypeID: string;
  public contentTypeLevel: number;
  public spContext: WebPartContext;
  constructor(context: WebPartContext) {
    this.spContext = context;
  }

  public async populateContentTypeLevel(listId: string) {
    let contentTypeHierarchy: IHandbookContentType[] = [];
    try {
      let oSPDataOperations = new SPDataOperations(this.spContext);
      let handbookContentTypes = await oSPDataOperations.getHandbookContentTypes(listId);
      await Promise.all(handbookContentTypes.map(async (handbookContentType) => {
        let parentContentType = await oSPDataOperations.getParentContentType(listId, handbookContentType.Id.StringValue);
        let handbookContentTypeDetails: IHandbookContentType = {
          contentTypeID: handbookContentType.Id.StringValue,
          name: handbookContentType.Name,
          parentContentTypeID: parentContentType.Id.StringValue,
          parentContentTypeLength: parentContentType.Id.StringValue.length,
          contentTypeLevel: 0,
          description: handbookContentType.Description
        };
        contentTypeHierarchy.push(handbookContentTypeDetails);
      }));

      let sortedObject = contentTypeHierarchy.sort((a, b) => (a.parentContentTypeLength > b.parentContentTypeLength) ? 1 : -1);
      let orderedHandbookContentTypes: IHandbookContentType[] = [];
      let parentItemLength = sortedObject[0].parentContentTypeLength;
      let level = 1;
      sortedObject.forEach((element, index) => {
        if (parentItemLength !== element.parentContentTypeLength) {
          level = ++level;
          parentItemLength = element.parentContentTypeLength;
        }

        let handbookContentTypeDetails: IHandbookContentType = {
          contentTypeID: element.contentTypeID,
          name: element.name,
          parentContentTypeID: element.parentContentTypeID,
          parentContentTypeLength: element.parentContentTypeLength,
          contentTypeLevel: level,
          description: element.description
        };
        orderedHandbookContentTypes.push(handbookContentTypeDetails);
      });
      return orderedHandbookContentTypes;
    } catch (error) {
      console.log('HandbookContentType.populateContentTypeLevel: ' + error);
    }
  }

  public async getAllColumnsForContentType(listId: string, itemId: number): Promise<HandbookColumn[]> {
    return new Promise<HandbookColumn[]>(async (resolve) => {
      if (itemId) {
        try {
          let oSPDataOperations = new SPDataOperations(this.spContext);
          let allFields = await oSPDataOperations.getAllColumnsForContentType(listId, itemId);
          let queryFields: any[] = [];
          let expandFields: any[] = [];
          let fieldType = {};
          let fieldLabel = {};
          let dependentLookupInternalNames: any = [];
          allFields.map((field) => {
            if (field.DependentLookupInternalNames !== undefined) {
              field.DependentLookupInternalNames.map((value) => {
                dependentLookupInternalNames.push(value);
              });
            }
          });

          allFields.map((field) => {
            if (dependentLookupInternalNames.indexOf(field.InternalName) === -1) {
              fieldType[field.InternalName] = field.TypeAsString;
              fieldLabel[field.InternalName] = field.Title;
              if (field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti') {
                queryFields.push(field.InternalName + '/Title');
                queryFields.push(field.InternalName + '/EMail');
                if (expandFields.indexOf(field.InternalName) === -1) {
                  expandFields.push(field.InternalName);
                }
              } else if (field.TypeAsString === 'Lookup') {
                queryFields.push(field.InternalName + '/Id');
                if (expandFields.indexOf(field.InternalName) === -1) {
                  expandFields.push(field.InternalName);
                }
              } else {
                queryFields.push(field.InternalName);
              }
            }
          });
          let itemData = await oSPDataOperations.loadPageContentType(listId, itemId, queryFields, expandFields);

          let handbookColumn: HandbookColumn[] = [];

          allFields.forEach(element => {
            handbookColumn.push({ name: element.Title, columnType: element.TypeAsString, internalName: element.InternalName, value: itemData[0][element.InternalName] });
          });
          this.columns = handbookColumn;
          let contentTypedata = await oSPDataOperations.getContentType(listId, itemId);
          let contentTypeDetails = await this.populateContentTypeLevel(listId);
          let contentType = contentTypeDetails.filter((contentTypeDetail) => contentTypeDetail.contentTypeID === contentTypedata[0].StringId);

          this.name = contentTypedata[0].Name;
          this.description = contentTypedata[0].Description;
          this.contentTypeID = contentTypedata[0].StringId;
          this.contentTypeLevel = contentType[0].contentTypeLevel;
        } catch (error) {
          console.log('HandbookContentType.getAllColumnsForContentType: ' + error);
        }
      }
      resolve(this.columns);
    });
  }

}